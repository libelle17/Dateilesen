Attribute VB_Name = "ZielDBFunktionen"
Option Explicit

Declare Function Tüt& Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long)
Declare Function OpenClipboard& Lib "user32" (ByVal hwnd&)
Declare Function CloseClipboard& Lib "user32" ()
Declare Function GetClipboardData& Lib "user32" (ByVal wFormat&)
Declare Function GlobalLock& Lib "kernel32" (ByVal hMem&)
Declare Function GlobalUnlock& Lib "kernel32" (ByVal hMem&)
Declare Function lstrcpy& Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any)
Declare Function GetWindowsDirectory& Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer$, ByVal nSize&)


Const BreiteSp1% = 28
Const LEDatei$ = "Leistungsexport_Porto.txt"

Const HADBName$ = "haerzte"
Public Const vgbVerspätung% = 21 ' wie lange ins neue Quartal die alte Abrechnung geht
Public Verspätung$
'Public Const kvneu = "SELECT GROUP_CONCAT(DISTINCT nachname) haname, ort, CAST(left(bsnr,7) AS char) kvnu, CONCAT(left(bsnr,2),'/',mid(bsnr,3,5)) kvnr, replace(tel1.tel,'-','') tel1, replace(fax,'-','') fax1, IF(obweibl,'Frau','Herr') anrede, titel, vorname, nachname, " & _
                     "if(ISNULL((SELECT genehmigung FROM " & hadbname & ".arzt_has_genehmigung ahg1 LEFT JOIN " & hadbname & ".genehmigung g1 ON g1.idgenehmigung = ahg1.genehmigung_id WHERE ahg1.arzt_id = idarzt AND genehmigung = 'DMP-DM1_Koordinierender Arzt_Hausarzt')),'0','1') dmpt1, " & _
                     "if(ISNULL((SELECT genehmigung FROM " & hadbname & ".arzt_has_genehmigung ahg2 LEFT JOIN " & hadbname & ".genehmigung g2 ON g2.idgenehmigung = ahg2.genehmigung_id WHERE ahg2.arzt_id = idarzt AND genehmigung = 'DMP-DM2_Koordinierender Arzt')),'0','1') dmpt2, " & _
                     "lanr, CONCAT(straße,' ',hausnr) straße, plz " & _
                     "FROM " & hadbname & ".bs " & _
                     "LEFT JOIN " & hadbname & ".ort ON ort_id = idort " & _
                     "LEFT JOIN " & hadbname & ".tel tel1 ON tel1.bs_id = idbs " & _
                     "LEFT JOIN " & hadbname & ".fax ON fax.bs_id = idbs " & _
                     "LEFT JOIN " & hadbname & ".arzt_has_bs ahb ON ahb.bs_id = idbs " & _
                     "LEFT JOIN " & hadbname & ".arzt a ON a.idarzt = ahb.arzt_id " & _
                     "LEFT JOIN " & hadbname & ".titel t ON t.idtitel = a.titel_id " & _
                     "GROUP BY kvnu"
'Public Const kvneu1 = "SELECT GROUP_CONCAT(DISTINCT nachname) haname, ort, CAST(left(bsnr,7) AS char) kvnu, " & _
                      "if(ISNULL((SELECT genehmigung FROM " & hadbname & ".arzt_has_genehmigung ahg1 LEFT JOIN " & hadbname & ".genehmigung g1 ON g1.idgenehmigung = ahg1.genehmigung_id WHERE ahg1.arzt_id = idarzt AND genehmigung = 'DMP-DM1_Koordinierender Arzt_Hausarzt')),'0','1') dmpt1, " & _
                      "if(ISNULL((SELECT genehmigung FROM " & hadbname & ".arzt_has_genehmigung ahg2 LEFT JOIN " & hadbname & ".genehmigung g2 ON g2.idgenehmigung = ahg2.genehmigung_id WHERE ahg2.arzt_id = idarzt AND genehmigung = 'DMP-DM2_Koordinierender Arzt')),'0','1') dmpt2 " & _
                      "FROM " & hadbname & ".bs " & _
                      "LEFT JOIN " & hadbname & ".ort ON ort_id = idort " & _
                      "LEFT JOIN " & hadbname & ".arzt_has_bs ahb ON ahb.bs_id = idbs " & _
                      "LEFT JOIN " & hadbname & ".arzt a ON a.idarzt = ahb.arzt_id " & _
                      "GROUP BY kvnu"
                      
' 11.9.15: kommt offenbar nicht vor, müsste ggf. nocht  DBCn.Execute "SET GROUP_CONCAT_MAX_LEN = 70" vorangestellt werden
Public Const kvneu2 = "SELECT GROUP_CONCAT(DISTINCT nachname) haname, ort, CAST(left(bsnr,7) AS char) kvnu, CONCAT(left(bsnr,2),'/',mid(bsnr,3,5)) kvnr, replace(tel1.tel,'-','') tel1, replace(tel2.tel,'-','') tel2, replace(fax1.fax,'-','') fax1, replace(fax2.fax,'-','') fax2, mail1.mail email, IF(obweibl,'Frau','Herr') anrede, titel, vorname, nachname, " & _
                      "MAX(if(ISNULL((SELECT genehmigung FROM " & HADBName & ".arzt_has_genehmigung ahg1 LEFT JOIN " & HADBName & ".genehmigung g1 ON g1.idgenehmigung = ahg1.genehmigung_id WHERE ahg1.arzt_id = idarzt AND genehmigung = 'DMP-DM1_Koordinierender Arzt_Hausarzt')),'0','1')) dmpt1, " & _
                      "MAX(if(ISNULL((SELECT genehmigung FROM " & HADBName & ".arzt_has_genehmigung ahg2 LEFT JOIN " & HADBName & ".genehmigung g2 ON g2.idgenehmigung = ahg2.genehmigung_id WHERE ahg2.arzt_id = idarzt AND genehmigung = 'DMP-DM2_Koordinierender Arzt')),'0','1')) dmpt2, " & _
                      "lanr, CONCAT(straße,' ',hausnr) straße, plz, GROUP_CONCAT(DISTINCT fachrichtung) zulg, Niederlassungsart Arzttyp, aktzeit " & _
                      "FROM " & HADBName & ".bs " & _
                      "LEFT JOIN " & HADBName & ".ort ON ort_id = idort " & _
                      "LEFT JOIN " & HADBName & ".tel tel1 ON tel1.bs_id = idbs " & _
                      "LEFT JOIN " & HADBName & ".tel tel2 ON tel2.bs_id = idbs AND tel2.tel <> tel1.tel " & _
                      "LEFT JOIN " & HADBName & ".fax fax1 ON fax1.bs_id = idbs " & _
                      "LEFT JOIN " & HADBName & ".fax fax2 ON fax2.bs_id = idbs AND fax2.fax <> fax1.fax " & _
                      "LEFT JOIN " & HADBName & ".mail mail1 ON mail1.bs_id = idbs " & _
                      "LEFT JOIN " & HADBName & ".arzt_has_bs ahb ON ahb.bs_id = idbs " & _
                      "LEFT JOIN " & HADBName & ".arzt a ON a.idarzt = ahb.arzt_id " & _
                      "LEFT JOIN " & HADBName & ".titel t ON t.idtitel = a.titel_id " & _
                      "LEFT JOIN " & HADBName & ".arzt_has_fachrichtung ahf ON ahf.arzt_id = a.idarzt " & _
                      "LEFT JOIN " & HADBName & ".fachrichtung fr ON fr.idfachrichtung = ahf.fachrichtung_id " & _
                      "LEFT JOIN " & HADBName & ".nlart ON idnlart = a.nlart_id " & _
                      "GROUP BY kvnu"

Const obAdo% = 0
Public FaxSendDatei$, DMPVorlage$
'Const FaxSendDatei$ = uVerz & "test1.txt"
'Const DMPVorlage$ = uVerz & "vorlagen\DMP-Vorlage gemein.dot"
' Inhalte folgender Variablen werden von der Procedure MedPlanNr an die sie aufrufenden weitergegeben:
Public Med$(), Dos$(), MedNr%, MedZahl%

Public lBehDat As Date
' aus AbrechFehler

Public Const artSpezÄrzte$ = "'ni','gstel','gs','rz','ep','bga','tk','APK','wd','ah'"
' ni = Notiz intern
' gstel = Telefonat Gerald Schade
' gs = Gerald Schade
' rz = Reinhard Zebhauser
' ep = Esther Preuss
' bga = Benedikt Gallitz
' tk = Thomas Kothny
' APK = Thomas Kothny in DocExpert
' wd = Diane Wagner
' ah = Anna Hammerschmidt

Public Const artSpezBerat$ = artSpezÄrzte & ",'wr','jl','ga','ih','cr','tn','be','lf'"
' wr = Walburga Roßmeier
' jl = Juliane Lange
' ga = Greß Angelika
' ih = Ines Hoffmann
' cr = Cornelia Reindl
' tn = Tanja Nuber
' be = Bender Elena
' lf = Larissa Fuchs

Public Const artSpezMA$ = "'tst','ke','hz','ns','mh','ag','ph','pq','er','ds','st','eb','us','sn','vb','mip','mm','rc','ik','ks','sb','cb','th','sp','ir','as','sa','sta','eg','ans','mc','rb','mi','gr','bs','sf','fs','eo','cd','mk'"
' tst = Tamara Sturm
' cr = Cornelia Reindl
'' eb = Elmas Balkan / Gürbüz
'' ke = Kupinic Elvisa
'' hz = Hilda Ziegler
'' ns = unbek. Mitarbeiter/Bedeutung? (2010, ein Eintrag 2012)
'' mh = Michaela Hofer
'' ag = Alexandra Gasse
'' ph = Petra Hermann
'' pq = Petra Quotschalla
'' er = Elke Rupprecht
' ds = Diana Sideris
' us = Uli Simon
'' st = Sahra Teherani
'' sn = Schilcher Nadine
'' vb = Veronika Blum
'' mm = Maria Mäusbacher
'' rc = Raffaela Corrao
'' ik = Ivonne Koller
'' ks = Kristina Scheible
' sb = Susanne Bittner
'' cb = Carmen Buffi
'' th = Tamara Hartmann
' sp = Silvia Plenagl
' mip = Michaela Pistek
'' ir = Ingrid Rymas
' AS = Stanka Andrasova
' sta = Stanka Andrasova
' sa = Alexandra Schley
' eg = Elmas Gürbüs
' ans = Angela Schmidt
' mc = Marina Chorosopoulou
' rb = Rafaela Bandula
' mi = Mina Muminovic
' gr = Gib Rabl
' bs = Blerta Sallteku
' sf = Sengpiel Franka
' fs = Franka Schweizer
' eo = Enkhmaa Oyunchimeg
' cd = Claudia Dannert
' mk = Melanie Kunze
'
' Einträge, die nicht automatisch mit einer Organuntersuchung verbunden sind
Public Const artSpezEintr$ = "'notiz','mbf','telef'," & artSpezBerat & "," & artSpezMA & _
",'pa','fa','fam','familie','beruf','sem','seminar','kind','bz','rp','B-rp','mus','muster','uzu','uz','hypo','colo','coloauf','haut','kard','aug','auge','augen'" & _
",'beweg','bewegung','bew','bewg','bewe','mess','ictauf','auto','insgd','vkgd','vkgd2'" & _
",'pros','prost','prostata','uro','impf','imp','gyn','caro','carotis'" & _
",'ap','mu','rauch','rauchen','alko','fams','schula','ass','kra','proc','au','GPD','GDP','ba','bla','kva','aufgd','HbA1cfre'" & _
",'ARCHIE2','ARCHIE','PATNRALPHA','arch','LZ','FDIAB1','FDIAB2','BLANKO','NV','M-HI','REZ-K','REZ-X','EDIAB1','LAS','EDIAB2','lab','Novo Rapid','PBA','PRAXG','FT','REZ-G','N','KK','REZ-P','anam','EKG01','KH','HV13','M-HM','L','IE','HÄK'" & _
",'tv','re','schul','schulung','50g','dak','dakap','dakluts','daknp','dmperg','unin','c19i','cia'"
' notiz = Texteintrag bis 12/06, seither "Merkblatt Fußsyndrom mitgegeben"
' Merkblatt Fußsyndrom Schley
' telef = Telefonat bis 4/06
' pa = Patientenannahme
' fa, fam, familie = Familienanamnese
' beruf = Beruf
' sem, seminar = Seminar
' kind = Kinderzahl
' bz = Blutzucker
' rp = Rezept
' B-rp = BTM-Rezept
' mus,muster = Ärztemusterabgabe
' uzu,uz = Unterzuckeranamnese
' hypo = Bericht über schwere Hypoglykämie mit Fremdhilfe oder Bewußtlosigkeit
' colo = Abfrage letzte Darmkrebsfrüherkennung
' coloauf = Anamnese Coloskopie und Aufklärung Darmkrebsfrüherkennung
' haut = Anamnese Hautscreening
' kard = Anamnese Kardiologe
' aug,auge,augen = Anamnese Augenarztbesuch bzw. Augenarztuntersuchung / -befund
' beweg, bewegung, bew, bewg, bewe = aktuelle Bewegungsanamnese
' mess = Messhäufigkeit (Makro)
' ictauf = Aufklärung über ICT-Aufklärung
' auto = Aufklärung über Insulin und Autofahren
' insgd = Insulinaufklärung bei Gestationsdiabetes
' vkgd = Verlaufskontrolle Gestationsdiabetes
' vkgd2 = Verlaufskontrolle Gestationsdiabetes ab 14.4.20
' pros, prost, prostata,uro = Anamnese Prostatakrebsvorsorge
' impf,imp = Anamnese Grippe- und Pneumokokkenimpfung
' gyn =  Anamnese Gynäkologenbesuch
' caro, carotis = Carotisdopplera-namnese
' ap = Bericht über Angina pectoris anamnestisch
' beruf = Beruf
' mu = Musterabgabe
' rauch, rauchen = aktueller Raucherstatus
' alko = aktuelle Alkoholgewohnheiten
' fams = Familienstatus
' schula = Schulungsanamnese
' ass = Vorteile und Risiken einer Thrombozytenhemmung besprochen
' kra = Bericht über diabetesbedingte Krankenhauseinweisung
' proc = Procedere
' au = Eintrag mit Infos über AU (au selbst hat andere Typ-Spalte)
' GPD,GDP = Auskunft über Verbleib/ Aktualisierung des Gesundheitspasses Diabetes
' ba, bla = Blutabnahme
' aufgd = Aufklärung Gestationsdiabetes
' HbA1cfre = HbA1c fremd
' ARCHIE2, ARCHIE = Archiveintrag in DocExpert
' PATNRALPHA,arch = weiterer Archiveintrag in DocExpert
' LZ = Leistungsziffer DocExpert
' FDIAB1,FDIAB2 = Abrechnung DMP DocExpert
' BLANKO, NV, M-HI,REZ-K,REZ-X,EDIAB1, LAS, EDIAB2,lab,Novo Rapid,PBA,PRAXG,FT,REZ-G,N,KK,REZ-P,anam,EKG01,KH,HV13,M-HM,L,IE,HÄK
' tv = Termin versäumt
' re = Rezept erstellt
' schul,schulung = Schulung
' 50g = 50g-Glucosetest vorher bei Gynäkologen durchgeführt
' dak = DAK-Programm-Makro
' dakap = DAK- Angiopathie
' dakluts = DAK-lower urinary tract-syndrome
' daknp = DAK-Neuropathie

' Einträge mit Untersuchungen, die vorne im Patientenlaufzettel stehen
Public Const artSpezUS0$ = "'gewicht','gewi','gew','rrvgl','rrvgln','bzvgl','bzvgln','fuß','füße','taille','urin'"
' gewicht, gewi, gew = Gewicht
' rrvgl = RR-Vergleich
' rrvgln = RR-Vergleich nicht möglich, Gerät nicht dabei
' bzvgl = Blutzuckervergleich
' bzvgln = Blutzuckervergleich nicht möglich, Gerät nicht dabei
' fuß,füße = genaue Fußinspektion
' taille = Taillenumfang
' urin = Urin-Teststreifenergebnis

' Einträge mit Untersuchungen, die nicht vorne im Patientenlaufzettel stehen
Public Const artSpezUS1$ = "'bzm','bztp','bks','anal','andm','andm2','angd','angd2','usal','usdm','usdm1','usdm2','bauch','ufrag','uabfrag'" & _
",'cgma','doppler','duplex','belastun'" & _
",'sono','sd','UKG','Größe','groe','HbA1c','hyper','keto','wv','vw','kv','komp','kompr','debr','ulc','ulcus','EKG','LZRR','Lufu','gpt'" & _
",'lactoset','trop','temp','oGTT','bmi','hüfte','rr','puls','GDT','bef','OAU','OA','inj','inf','ht','htt','ADL','TUG','who','vac','vacc','grippe'"
' bzm = Blutzuckermessung
' bztp = Blutzuckertagesprofil
' bks = Blutkörperchensenkungsgeschwindigkeit
' anal = Anamnese allgemein
' andm = Anamnese Diabetes mellitus
' andm2 = Anamnese Diabetes mellitus ab 13.4.20
' angd = Anamnese Gestationsdiabetes
' angd2 = Anamnese V.a. Gestationsdiabetes
' usal = Untersuchung allgemein
' usdm, usdm1, usdm2 = Untersuchung Diabetes mellitus
' bauch = Bauchuntersuchung
' ufrag, uabfrag = Anamnese Urinsymptome
' cgma = CGM-Anlage
' doppler = Dopplerbefund bei Untersuchung mit Stiftsonde
' duplex = Duplexbefund
' belastun = Belastungsuntersuchung
' sono = Sonobefund
' sd = Sonobefund Schilddrüse
' UKG = Echokardiogramm
' Größe,groe = Körpergröße
' HbA1c = HbA1c-Messung bei uns
' hyper = Hyperglykämien mit Fremdhilfen
' keto = Ketoazidosen
' wv = Wundverband, vw = Verbandswechsel
' kv,komp,kompr = Kompressionsverband
' debr = Debridement
' ulc, ulcus = Ulcusbeschreibung
' EKG = Befunddaten EKG
' LZRR = Befunddaten Langzeitblutdruckmessung
' Lufu = Befunddaten Lufu
' gpt  = Gastroparesetest
' lactoset = Lactose-Toleranztest
' trop = Troponintest
' temp = Temperatur
' oGTT = oraler Glucosetoleranztest
' bmi = BMI
' hüfte = Hüftumfang
' rr = Blutdruck
' puls = Puls
' GDT = Befunddaten customed-Untersuchung
' bef = Größe und Gewicht in DocExpert
' OAU,OA = Oberarmumfang
' inj = Injektion
' inf = Infusion
' ht,htt = Hochtontherapie
' ADL =
' TUG =
' who = WHO-5-Depressionsbogen
' vac,vacc,grippe = Impfung

' Einträge, die mit einer Organuntersuchung verbunden sind
Public Const artSpezUS$ = artSpezUS0 & "," & artSpezUS1

' für Liste 112
Public Const artSpezSonst$ = "'PKon','kva','rcg','dd','dg','ddg','bddg','rdg','lar','plar','spbrp','htxt','txt','priv','bez','rech'" & _
",'pvs','mhng','Ch.B.','prp','unbek','WZeit','ict','ct','oad','typ1','BfA','PGeb','haem','Text','kgä','kgp','ksä','ksp','kmä','kmp'" & _
",'mimp','uebw','abstr','cib','nfd','dpe'"
' Pkon = Behandlungsdauer
' kva = Abrechnungsziffer gesetzliche Kasse
' rcg = Rechnungsnummer
' dd = Texteintrag gelöschte Dauerdiagnose
' dg = Texteintrag gelöschte Einzeldiagnose
' ddg = Texteintrag gelöschte Dauerdiagnose
' bddg = alte Diagnose?
' rdg = Rechnungsdiagnose
' lar = Langrezept
' plar = Ergänzungstext Langrezept
' spbrp = Ergänzungstext Langrezept
' htxt = Text für Verordnung (Podologie)
' txt = Text für Verordnung (Podologie)
' priv = Privatabrechnungsziffer
' bez = Einbezahlung Privatrechnung
' rech = Auskunft über Einzahlung einer Privatrechnung
' pvs = Privatrechnung
' mhng = Mahnung
' Ch.B. = Chargenbezeichnung bei Fremdeiweiß / Impfung
' prp = Privatrezept
' unbek = importierte Daten unbekannter Art
' WZeit = Wartezeit
' ict = Warteliste Schulung ICT bei Pat. 1789 (Seminar)
' ct = Warteliste Schulung CT bei Pat. 1789 (Seminar)
' oad = Warteliste Schulung OAD bei Pat. 1789 (Seminar)
' typ1 = Warteliste Schulung Typ 1 bei Pat. 1789 (Seminar)
' BfA = Honorar für ärztlichen Befundbericht BfA
' PGeb = Praxisgebührmahnung gedruckt
' haem = Ergebnis Test auf okkultes Blut im Stuhl
' Text = Diagnosentext
' dmperg = DMP-Ergänzung
' mimp = Medimport-Datei
' uebew = altes Überweisungsformular?
' kgä = kein GLP-1-Analogon von ärztlicher Seite gewünscht
' kgp = kein GLP-1-Analogon von Patienten-Seite gewünscht
' ksä = kein SGLT-2-Hemmer von ärztlicher Seite gewünscht
' ksp = kein SGLT-2-Hemmer von Patienten-Seite gewünscht
' kmä = kein Metformin von ärztlicher Seite gewünscht
' kmp = kein Metformin von Patienten-Seite gewünscht
' unin = unerwartete Inanspruchnahme
' abstr = Abstrich
' cib = Covid-Impfberatung
' nfd = Notfalldatensatz
' dpe = Datensatz persönliche Erklärung

Public Const artspezG$ = artSpezEintr & "," & artSpezUS
Public Const AuffArtSql = vbCrLf & _
 "SELECT art, Pat_ID, gesname(pat_id) Name, Zeitpunkt, Inhalt FROM eintraege " & vbCrLf & _
 "WHERE NOT art IN (" & artspezG & "," & artSpezSonst & ") " & vbCrLf & _
 "AND zeitpunkt > 20161231 " & vbCrLf & _
 "ORDER BY zeitpunkt DESC; "
'Public Const artinartspez$ = "(art in " & artSpez
Public sql$, sql1$ ' SQL-Text für alle möglichen Abfragen
Dim QMdbAkt$, nzw$
Dim DmPStrS$
Dim Pat_id& ' für dii(
Public Enum DFSNiveau
   stNichts%
   St0%
   St1%
   St2%
   AmpZeh%
   St3%
   AmpUS%
   St4%
   AmpOS%
  End Enum
Public Enum ZigSt
 nie%
 früher%
 aktuell%
 vorlangem%
End Enum
Public Enum TherapieArt
 offen
 Diät
 OAD
 glp1
 komb
 glp1ins
 ct
 glp1ict
 ict
 csii
End Enum
'Dim rDT AS DAO.Recordset
'Dim rFot AS DAO.Recordset ' solange Fotos noch in Access-Datenbank
Dim SStr$
Public Const Titel$ = "KV-Ärzte raussuchen lassen"
Public WinDir$
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096


Public lDMPPat_id&
Public Const forminhalt$ = "SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, form_abk FROM (((`forminhfeld` LEFT JOIN `forminhkopf` ON `forminhfeld`.foid=`forminhkopf`.foid) LEFT JOIN `formulare` ON `formulare`.formid=`forminhkopf`.form_id) LEFT JOIN `forminhaltfeld` ON `forminhfeld`.feldvw=`forminhaltfeld`.feldvw) LEFT JOIN `forminhaltfeldinh` ON `forminhfeld`.feldinhvw=`forminhaltfeldinh`.feldinhvw "
'Public LeistBDT$
Public Enum DMPStat
 unauff
 ndok
 auff
 pathdok ' als pathologisch dokumentiert, nur bei sens
End Enum
Public Enum DMPSeite
 unbek
 gleich
 rE
 li
End Enum
Public Enum DMPEmpf
 halten
 senken
End Enum
Public Enum DMPaug
 durchg
 ndurch
 veranl
End Enum
Public Enum AntidiabMedType
 adja = -1
 adnein = 0
 adki
End Enum

Public Type DMPClass
 Pat_id As Long
 Nachname As String
 Vorname As String
 NVorsatz As String
 NVors As String
 GebDat As Date
 Titel As String
 strasse As String
 Hausnr As String
 Versichertennummer As String * 13
 VKNr As String * 5
 plz As String * 5
 ort As String
 Lkz As String * 5
 PFPlz As String
 PFOrt As String
 Geschlecht As String
 Anschrzus As String
 Postleitzahl As String
 PrivatTel As String
 KVKStatus As String * 1
 IK As String * 9 '7 => lieber 9 s. Wiki
 VschBeg As Date
 KtrAbrB As String * 2
 dmpVertret As String * 1
 dmpArztw As String * 1
 dmpHypos As String * 2
 dmpKhsA As String * 2
 dmpDMSchulEmpf As String
 dmpDMSchulWahrg As String
 dmpHypertSchulEmpf As String
 dmpHypertSchulWahrg As String
 dmpKKTabakEmpf As String * 1
 dmpKKErnEmpf As String * 1
 dmpKKkTrainEmpf As String * 1
 dmpHbA1cZiel As String * 4
 dmpUewFuss As String * 1
 dmpEinwDM As String * 1
 dmphalbj As String * 1
 
 dtyp As String ' Diabetestyp, "1", "2"
 daseit As String
 dspsy As Boolean ' diabetesspezifische Symptome
 dspmed As Boolean ' diagnosespezifische Medikation
 FE(17) As Integer ' Folgeerkrankungen bekannt
 FEn(17) As Integer ' Folgeerkrankungen
 FEi(17) As String ' ICD der Folgeerkrankung
 FEd(17) As Date ' Datum der ICD der Folgeerkrankung
 SE(2) As Date ' Schulungen empfohlen
 kSE As Date ' bei letzter Doku keine Schulung empfohlen
 SW(5) As Date ' Schulungen wahrgenommen ' (0) = D.m., (1) = Hypertonie, (2) = keine
 DS(2) As Date ' DiabSchulung
 HS(2) As Date ' HyperSchulung
 VorM(13) As Integer
 obGlib As AntidiabMedType
 obmetf As AntidiabMedType
 obGlucI As AntidiabMedType
 obSHGlin As AntidiabMedType
 obGlit As AntidiabMedType
 obdpp4 As AntidiabMedType
 obglp1 As AntidiabMedType
 obsglt2 As AntidiabMedType
 obSonstAD As AntidiabMedType
 insz As Integer
 obIns As Boolean
 obAnal As Boolean
 obHMG As Boolean
 obAntihyp As Boolean
 obACEH As Boolean
 obAT1 As Integer
 obBetabl As Boolean
 obThro As Boolean
 obOAK As Integer
 obDiur As Integer
 Tabak As Boolean
 kgr As Integer ' Körpergröße
 gewi As Single ' Gewicht
 bmi As Single ' Body Mass Index
 Puls As DMPStat ' Pulsstatus
 sens As DMPStat ' Sensibiliätsprüfung
 PrRR As String ' Blutdruck
 RRsyst As Integer
 RRdiast As Integer
 bekHb As Single ' HbA1c
 Crea As Single ' Kreatinin
 eGFR As Single ' eGFR
 fußst As DMPStat ' Fußstatus
 maxAmp As DFSNiveau ' nur intern
 NStSeite As DMPSeite ' nur intern
 oap As DMPSeite ' Osteoarthropathie; 'gleich' = unbekannt, 'unbek' = keine
 mSei As DMPSeite ' schwerer betroffener Fuß
 mWA As String ' maximales Wagner-Stadium
 khew As Boolean ' Krankenhauseinweisung
 mau As DMPStat ' Mikroalbuminurie
 oblaser As DMPStat ' Lasertherapie
 hypoZAn As Integer ' Zahl der Hypos in der Anamnese
 hypoZKK As Integer ' Zahl der Hypos in Karteikarte
 hypoZ As Integer ' Zahl der Hypos, Synthese
 krZAn As Integer ' Zahl der stationären Aufenthalte wegen Diabetes aus Anamnese
 krZKK As Integer ' Zahl der stationären Aufenthalte wegen Diabetes aus Karteikarte
 krz As Integer ' Zahl der Khs-Aufenthalte, Synthese
 obVorb As Boolean ' ob Vorbefund da
 obNI As Boolean ' ob Niereninsuffizienz
 neuDial As Boolean ' ob neu Dialyse
 neuErbl As Boolean ' ob neu Erblindung
 neuAmp As Boolean ' ob neu Amputation
 neuMI As Boolean ' ob neu Myokardinfarkt
 neuApo As Boolean ' ob neu Apoplex
 ernb As Boolean ' Ernährungsberatung
 hbEmpf As DMPEmpf ' HbA1c-Empfehlung
 rrEmpf As DMPEmpf ' Blutdruckempfehlung
 aug As DMPaug ' Augenuntersuchung
 tart As TherapieArt ' Therapieart
 obSchw As Boolean
 obG632 As Boolean ' periphere sensomotorische Neuropathie
 obRauch As Boolean ' ob Patient raucht (und deshalb ein Angebot von der Kasse zum Rauchen aufhören wünscht
 obBeweg As Boolean ' ob der Patient ein Bewegungsangebot von der Kasse haben will
 inj As String * 1 ' Injektionsstellen
 ulcus As Integer '0=oberflächlich, 1=tief, 2=nein, 3=nicht untersucht
 NaeUs As Integer ' 0=1 Jahr, 1=6 Monate, 2=3 Monate
 Infekt As Integer '0=nicht untersucht, 1=ja, 2=nein
 Deform As Integer '0=nein, 1=ja
 Hyperk As Integer '0=nein, 1=ja
 ZnUlcus As Integer '0=nein, 1=ja
 ZnAmput As Integer '0=nein, 1=ja
End Type

Enum LabArt
 LabArt0
 LA_HbA1c
 LA_HbA1cmm
 la_krea
 LA_AlbCre
 LA_Kali
 LA_Chol
 LA_Hdl
 LA_LDL
 LA_eGFR
 LA_TG
 LA_GGT
 LA_GPT
 LA_FERR
 LA_DIGO
 LA_DIGI
 LA_TSH
 LA_fT4
 LA_ft3
 LA_Hb
 LA_B12
 LA_FOL
 LA_LEUK
 LA_CRP
 LA_CK
End Enum ' LabArt
' statische Variable für die Funktion LabPat
Public alt_la As LabArt, altPID&, aktlwx&, lwZahl&
Public Lab() As labtyp

'#Const sqllangsam = True
#If sqllangsam Then
' nach Vorbild von Public Function LabPat(la AS LabArt, Pat_id&) AS labtyp
Public Function LetztLab(pid&, Abkü$, Einh$, Zp As Date) As labtyp
 Dim gefunden%, i&
 On Error GoTo fehler
 ' bei neuem Patienten sein ganzes Labor einlesen
 If pid <> altPID Then
  Dim rsl As New ADODB.Recordset ' Labor
  Set rsl = hollabor(PatID:=pid, Abkü:="", zpkl:=0, wertkl:=0, obnachgruppe:=0, Zahl:=lwZahl)
  If Not rsl.BOF And lwZahl Then
   ReDim Lab(lwZahl)
   i = 0
   Do While Not rsl.EOF
    Lab(i).Abkü = rsl!Abkü
    Lab(i).WertSg = Replace$(rsl!Wert, ".", ",")
    Lab(i).Einheit = rsl!Einheit
    Lab(i).Zp = rsl!Zeitpunkt
    i = i + 1
    rsl.MoveNext
   Loop
  End If
  Set rsl = Nothing
 End If ' pid <> altPID Then
 gefunden = 0
 If lwZahl Then
  For i = 0 To lwZahl
   If Lab(i).Abkü = Abkü And Lab(i).Einheit = Einh And Lab(i).Zp < Zp Then
    LetztLab.Abkü = Lab(i).Abkü
    LetztLab.Einheit = Lab(i).Einheit
    LetztLab.WertSg = Replace$(Lab(i).WertSg, ".", ",")
    LetztLab.Zp = Lab(i).Zp
    gefunden = True
    Exit For
   End If
  Next i
  If Not gefunden Then
   LetztLab.Abkü = vNS
   LetztLab.Einheit = vNS
   LetztLab.WertSg = vNS
   LetztLab.Zp = -1
  End If
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LetztLab/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
#End If


' wenn nichts (mehr) gefunden, dann wird zp auf 0 gesetzt
Public Function LabPat(la As LabArt, pid&, Optional naechster%) As labtyp
 Dim gefunden%
 On Error GoTo fehler
 ' bei neuem Patienten sein ganzes Labor einlesen
 If pid <> altPID Then
  Dim rsl As New ADODB.Recordset ' Labor
  Dim i&
  Set rsl = hollabor(pid, "", 0, 0, 0, lwZahl)
  If Not rsl.BOF And lwZahl Then
   ReDim Lab(lwZahl)
   i = 0
   Do While Not rsl.EOF
    Lab(i).Abkü = rsl!Abkü
    Lab(i).WertSg = Replace$(rsl!Wert, ".", ",")
    Lab(i).Einheit = rsl!Einheit
    Lab(i).Zp = rsl!Zeitpunkt
    i = i + 1
    rsl.MoveNext
   Loop
  End If
  Set rsl = Nothing
 End If
 ' dann entscheiden, ob mit dem Satzzeiger von vorne angefangen werden muss
 Dim angefx&
 If pid <> altPID Or alt_la <> la Then
  angefx = 0
  aktlwx = 0
 Else
  angefx = True
  If Not naechster Then aktlwx = aktlwx - 1
 End If
 gefunden = 0
 If lwZahl Then
  Do While True
   If angefx Then aktlwx = aktlwx + 1 ' der 1. Datensatz könnte schon richtig sein
   If aktlwx > lwZahl Then Exit Do
'   Debug.Print Lab(aktlwx).Abkü
   If Lab(aktlwx).WertSg <> "" Then
    If obLabI(la, Lab(aktlwx)) Then
     gefunden = True
     Exit Do
    End If
   End If
   angefx = True
  Loop
  If gefunden Then
   LabPat.Abkü = Lab(aktlwx).Abkü
   LabPat.Einheit = Lab(aktlwx).Einheit
   LabPat.WertSg = Replace$(Lab(aktlwx).WertSg, ".", ",")
   LabPat.Zp = Lab(aktlwx).Zp
  Else
   LabPat.Abkü = vNS
   LabPat.Einheit = vNS
   LabPat.WertSg = vNS
   LabPat.Zp = -1
  End If
 Else
  LabPat.Zp = -1
 End If ' lwzahl
 altPID = pid
 alt_la = la
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabPat/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LabPat(la AS LabArt, pid&) AS labtyp

'Sub testlabpat()
' Call Lese.ProgStart
' Dim la AS labtyp
' la = LabPat(LA_Krea, 14)
' Debug.Print la.Wert & " " & la.Abkü & " " & la.Zp & " " & la.Einheit
'End Sub
'Sub testlabp()
' Call Lese.ProgStart
' Dim la AS labtyp
' la = LabPat(LA_eGFR, 14)
' Debug.Print la.Wert & " " & la.Abkü & " " & la.Zp & " " & la.Einheit
'End Sub

'Function LabEPat(la AS LabArt, Pat_id&) AS ADODB.Recordset
' Set LabEPat = New ADODB.Recordset
' LabEPat.Open LabEPatS(la, Pat_id).Value, DBCn, adOpenStatic, adLockReadOnly
'End Function ' LabEPat
Function obLabI%(la As LabArt, LT As labtyp) ' ob Labor interessant
 Select Case la
  Case LA_HbA1c:
   obLabI = (LT.Abkü Like "HBA[C1]*" And MachNumerisch(LT.WertSg) < 22) ' 30.8.14: wegen falsch eingetragener Fremdlabore gestrichen: AND `einheit` = '%'
            ' identisch, wg.Kompatibilität nicht genommen: "rlike 'hba[c1]' AND einheit = '%'"
  Case LA_HbA1cmm:
   obLabI = (UCase(LT.Abkü) Like "HBA[C1]*" And MachNumerisch(LT.WertSg) >= 22) ' 30.8.14: wegen falsch eingetragener Fremdlabore gestrichen: AND `einheit` = '%'
  Case la_krea:
  ' abkü in ('creat','krea02','krea','krea02','kres') AND einheit='mg/dl'
   obLabI = ((LT.Abkü = "CREAT" Or LT.Abkü = "KRE02" Or LT.Abkü = "KREA" Or LT.Abkü = "KREA02" Or LT.Abkü = "KRES") And LT.Einheit = "mg/dl")
  Case LA_AlbCre:
  ' zu koordinieren mit doViewserstellen: obnephrop
   obLabI = ((LT.Abkü = "ALBCRE" Or LT.Abkü = "ALBKRE" Or LT.Abkü = "ALBQ" Or LT.Abkü = "ALBUM" Or LT.Abkü = "ALBUP") And LT.Einheit Like "mg/g *") Or ((LT.Abkü = "ALBU" Or LT.Abkü = "ALBUMU") And (LT.Einheit = "mg/l" Or LT.Einheit = ""))
  Case LA_Kali:
   obLabI = (LT.Abkü = "K" Or LT.Abkü = "KALI")
  Case LA_Chol:
   obLabI = (LT.Abkü = "CHOL")
  Case LA_Hdl:
   obLabI = (LT.Abkü = "HDL" Or LT.Abkü = "HDLC")
  Case LA_LDL:
   obLabI = (LT.Abkü Like "LDL*" And LT.Einheit = "mg/dl")
  Case LA_eGFR:
   obLabI = (LT.Abkü Like "GFR*" Or UCase(LT.Abkü) Like "GFC*" Or LT.Abkü Like "CREACL*" Or LT.Abkü = "MDRD") ' GFCK4
   ' "GFR','GFRCYS','GFRK','GFRM','GFRT','GFRW','CREACL','GFC','GFCK1','GFCK2','GFCK3','GFCK4','GFCM1','GFCW1','GFCW2','MDRD"
  Case LA_TG:
   obLabI = (LT.Abkü = "TRI" Or LT.Abkü = "TRIG" Or LT.Abkü = "NTFE")
  Case LA_GGT:
   obLabI = (LT.Abkü = "GAGT" Or LT.Abkü = "GAGT02" Or LT.Abkü = "GGT" Or LT.Abkü = "GT" Or LT.Abkü = "YGT")
  Case LA_GPT:
   obLabI = (LT.Abkü = "GPT" Or LT.Abkü = "GPTR02" Or LT.Abkü = "GPTR03")
  Case LA_FERR:
   obLabI = (LT.Abkü = "FERR" Or LT.Abkü = "FERR01" Or LT.Abkü = "FERRI" Or LT.Abkü = "FERRIT" Or LT.Abkü = "FERRO")
  Case LA_DIGO:
   obLabI = (LT.Abkü = "DIGO" Or LT.Abkü = "DIGOX")
  Case LA_DIGI:
   obLabI = (LT.Abkü = "DIGI" Or LT.Abkü = "DIGIT")
  Case LA_TSH:
   obLabI = (LT.Abkü = "TSH" Or LT.Abkü = "TSHE" Or LT.Abkü = "TSH-L_K" Or LT.Abkü = "TSBL" Or LT.Abkü = "TSBF" Or LT.Abkü = "TSH2" Or LT.Abkü = "TSHTRH")
  Case LA_fT4:
   obLabI = (LT.Abkü Like "FT4*")
  Case LA_ft3:
   obLabI = (LT.Abkü Like "FT3*")
  Case LA_Hb:
   obLabI = (LT.Abkü = "HB")
  Case LA_FOL:
   obLabI = (LT.Abkü = "FOL" Or LT.Abkü = "FOL_S" Or LT.Abkü = "FOLS" Or LT.Abkü = "FOLS_K" Or LT.Abkü = "FOLS01" Or LT.Abkü = "FOLSN")
  Case LA_LEUK:
   obLabI = (LT.Abkü = "LEU" Or LT.Abkü = "LEUK" Or LT.Abkü = "LEUKO" Or LT.Abkü = "LEUK_H" Or LT.Abkü = "LEUK_T")
  Case LA_CRP:
   obLabI = (LT.Abkü = "CRP" Or LT.Abkü = "CRPF" Or LT.Abkü = "CRPHS" Or LT.Abkü = "CRPQ" Or LT.Abkü = "CRPQ01" Or LT.Abkü = "CRPST")
  Case LA_CK:
   obLabI = (LT.Abkü = "CK" Or LT.Abkü = "CKE" Or LT.Abkü = "CKNAC")
  Case LA_B12:
'   obLabI = (lt.Abkü = "B12" Or lt.Abkü = "B12N" Or lt.Abkü = "B12akt" Or lt.Abkü = "VB12" Or lt.Abkü = "VITB12")
   obLabI = (LT.Abkü = "B12" Or LT.Abkü = "VITB12" Or LT.Abkü = "B12N")
 End Select
End Function ' obLabI%(LA AS LabArt, lt AS labtyp) ' ob Labor interessant
'' nur noch in Niereninsuffizienzpauschalendiabetiker_Click
'Function LabEPatS(la AS LabArt, Pat_id&) AS CString ', Optional minZp$) AS CString
'#If bis0418 Then
' Const s1$ = "SELECT date(zeitpunkt) zp, IF(ISNULL(wert),if(ISNULL(kommentar),'',kommentar),wert) Wert FROM `laborneu` ln LEFT JOIN laborkommentar lk ON ln.kommentarvw = lk.kommentarvw WHERE "
' Const s2$ = " AND pat_id = "
' Const s3$ = " UNION SELECT date(u.eingang) zp, IF(ISNULL(w.wert),if(ISNULL(w.kommentar),'',w.kommentar),w.wert) Wert FROM `laborxus` u LEFT JOIN laborxwert w ON u.refnr = w.refnr WHERE "
'#End If
' Const s4$ = " GROUP BY zp ORDER BY zp DESC"
' Dim Abkü$
' SELECT CASE la
'  Case HbA1c: Abkü = "`abkü` in ('HBA1','HBA1C','HbA1c Eigenlabor','HBA1C_','HBAC') AND CAST(CONCAT('0',replace(`wert`,',','.')) AS decimal) < 22 " ' 30.8.14: wegen falsch eingetragener Fremdlabore gestrichen: AND `einheit` = '%'"
'            ' identisch, wg.Kompatibilität nicht genommen: "rlike 'hba[c1]' AND einheit = '%'"
'  Case HbA1cmm: Abkü = "`abkü` in ('HBA1','HBA1C','HbA1c Eigenlabor','HBA1C_','HBAC') AND CAST(`wert` AS decimal) >= 22 " ' 30.8.14: wegen falsch eingetragener Fremdlabore gestrichen: AND `einheit` = '%'"
'  Case Krea: Abkü = "abkü in ('CREAT','KRE02','KREA','KREA02','KRES') AND einheit = 'mg/dl'"
'  Case AlbCre: Abkü = "((abkü in ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND einheit like 'mg/g %') OR (abkü in ('ALBU','ALBUMU') AND (einheit = 'mg/l' OR einheit = '')))"
'  Case Kali: Abkü = "abkü in ('K','KALI')"
'  Case Chol: Abkü = "abkü = 'CHOL'"
'  Case Hdl:  Abkü = "abkü IN ('HDL','HDLC')"
'  Case LDL:  Abkü = "abkü like 'LDL%' AND einheit = 'mg/dl'" ' in ('LDL','LDLB','LDLH','LDLH01','LDLLG','LDLS')
'  Case eGFR: Abkü = "(abkü RLIKE '^gfr|^gfck|^creacl')"
' End SELECT
' Set LabEPatS = New CString
' 'If minZp = "" Then
'#If bis0418 Then
' LabEPatS.AppVar Array(s1, Abkü, s2, Pat_id, s3, Abkü, s2, Pat_id, s4)
'#End If
' 'Else
''    LabEPatS.AppVar Array(s1, Abkü, s2, Pat_id, " AND zeitpunkt >= ", minZp, s3, Abkü, s2, Pat_id, " AND u.eingang >= ", minZp, s4)
'' End If
'Const Werte$ = "pat_id, Zeitpunkt zp, abkü, abk_ur, Langtext,CONVERT(replace(CONCAT('0',Wert),',','.'),decimal(10,1)) Wert,Einheit,Einheit_ur,Kommentar,NB,NB_ur,uNg,uNg_ur,oNg,oNg_ur,Labor,Pfad "
' LabEPatS.AppVar Array("SELECT * FROM (SELECT ", Werte, " FROM `labor2a` WHERE pat_id = ", Pat_id, " AND ", Abkü, " UNION SELECT ", Werte, " FROM `labor1a` WHERE pat_id = ", Pat_id, " AND ", Abkü, " ) i", s4)
''  sql = "SELECT pat_id, zeitpunkt zp, wert FROM `laborneu` ln WHERE abkü RLIKE 'hba[c1]' AND einheit = '%' AND pat_id = " & Pat_id & " UNION SELECT u.pat_id, u.eingang zp, w.wert FROM `laborxus` u LEFT JOIN laborxwert w ON u.refnr = w.refnr WHERE abkü RLIKE 'hba[c1]' AND einheit = '%' AND pat_id = " & Pat_id & " ORDER BY pat_id,zp DESC"
'End Function ' LabEPatS

Function latinis$(artSpez$)
 latinis = Replace$(artSpez, ",'", ",_latin1'")
 If Left$(artSpez, 1) = "'" Then latinis = "_latin1" & latinis
End Function ' latinis$(artSpez$)

Public Sub rrpruef()
 Dim rs As New ADODB.Recordset
 Dim ad As DMPClass
 Dim raAna As New ADODB.Recordset
 Lese.ProgStart
 rs.Open "SELECT pat_id, gesname(pat_id) Name,(SELECT max(zeitpunkt)from rr WHERE pat_id=n.pat_id) rr1 FROM namen n ORDER BY pat_id DESC", DBCn, adOpenStatic, adLockReadOnly ' WHERE pat_id <60050
 While Not rs.EOF
  Debug.Print rs!Pat_id
'  Call DMPString$(rs!Pat_id, ad, , , nz(rs!RR1,Now()), 0)
  Set raAna = Nothing
  If obAdo Then
   raAna.Open "SELECT patAlter(ana.pat_id) PAlter, ana.* FROM `anamnesebogen` ana WHERE pat_id = " & rs!Pat_id, DBCn, adOpenKeyset, adLockReadOnly
  Else
   Call raAna.Open("SELECT patAlter(ana.pat_id) PAlter, ana.* FROM `anamnesebogen` ana WHERE pat_id = " & rs!Pat_id, DBCn, adOpenDynamic, adLockReadOnly)
  End If
  ad.RRsyst = 0
  ad.RRdiast = 0
  ad.PrRR = GetPrRR(raAna, ad.RRsyst, ad.RRdiast, obdiastkorr:=True)
  If ad.RRsyst > 300 Or ad.RRsyst < 50 Or ad.RRdiast > 180 Or ad.RRdiast < 30 Then If Not IsNull(rs!RR1) Then Stop
  ad.RRsyst = 0
  ad.RRdiast = 0
  rs.MoveNext
 Wend
 Debug.Print "Fertig mit rrpruef"
End Sub ' rrpruef

' in do_DMPAusgebStandAlone, dodoPLZ, DMPAusgeb0, doCallDMP
Function DMPString$(pid&, üdt As DMPClass, Optional ohneVorDMP%, Optional mitBezeich%, Optional DokuDat As Date, Optional mitStr% = True)
 Dim ErgebDatei$
 
 If Not DokuDat Then DokuDat = Min(Now(), QEnd(ZQuart(Now() - Verspätung)))
 ErgebDatei$ = aVerz & "DMP.txt"
' Dim aktDC.obglib AS AntidiabMedType, aktDC.obmetf AS Boolean, aktDC.obgluci AS Boolean, aktDC.obshglin AS Boolean, aktDC.obglit AS Boolean, aktDC.obsonstAD AS boolean, _
     aktDC.obins AS Boolean, aktDC.obanal AS Boolean, aktDC.obhmg AS Boolean, aktDC.obantihyp AS Boolean, aktDC.obthro AS Boolean
'Dim rEi AS DAO.Recordset, rLaU AS DAO.Recordset, rKH AS DAO.Recordset
Dim raEi As New ADODB.Recordset, raKH As New ADODB.Recordset
'Dim raLau AS New ADODB.Recordset, raLaum AS New ADODB.Recordset
Dim Labs As labtyp, LabSm As labtyp
Dim rUlc As New ADODB.Recordset
Dim rfuss As New ADODB.Recordset
Dim rfal As New ADODB.Recordset
Dim lapp As New ADODB.Recordset
'Dim ralauSql$
Dim aktDC As DMPClass
'Dim rdo AS DAO.Recordset
'Dim rado AS New ADODB.Recordset
Dim raAna As New ADODB.Recordset
Dim raDT As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rsAnam As New ADODB.Recordset
Dim GesNa$
Dim DMSchulz%, HSchulz%, SchulStr$
Dim Tbl
'Dim lsql$
'Dim RRsyst%, RRdiast%
Dim fiabfr$
'Dim obVorb% ' ob vorbefund da
' Beim ersten Aufruf von diI muß Pid enthalten sein
Dim ab315%
ab315 = (DokuDat > #6/30/2015#)
Dim ab317%
ab317 = (DokuDat > #6/30/2017#)
nochmal:
If diI("Z34 Z33", pid) Then aktDC.obSchw = True
If diI("E10") Then aktDC.dtyp = "1" Else If diI("E11") Then aktDC.dtyp = "2"
If Not ohneVorDMP Then
'fiabfr = "SELECT Pat_ID, FID, form_abk,Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, form_abk FROM (((`forminhfeld` LEFT JOIN `forminhkopf` ON `forminhfeld`.foid=`forminhkopf`.foid) LEFT JOIN `formulare` ON `formulare`.formid=`forminhkopf`.form_id) LEFT JOIN `forminhaltfeld` ON `forminhfeld`.feldvw=`forminhaltfeld`.feldvw) LEFT JOIN `forminhaltfeldinh` ON `forminhfeld`.feldinhvw=`forminhaltfeldinh`.feldinhvw WHERE form_abk like ""dmpdtyp%"" AND feldinh = ""X"" AND pat_id = " & Pid & " ORDER BY zeitpunkt"
fiabfr = "SELECT * FROM (SELECT Pat_ID, FID, form_abk, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh FROM `forminhkopf` LEFT JOIN `formulare` ON `formulare`.formid = `forminhkopf`.form_id LEFT JOIN `forminhfeld` ON `forminhfeld`.foid = `forminhkopf`.foid LEFT JOIN `forminhaltfeld` ON `forminhfeld`.feldvw=`forminhaltfeld`.feldvw LEFT JOIN `forminhaltfeldinh` ON `forminhfeld`.feldinhvw=`forminhaltfeldinh`.feldinhvw AND feldinh='X' WHERE pat_id = " & pid & " AND form_abk like 'dmpdtyp%') AS i WHERE feldinh<>null ORDER BY zeitpunkt"
raAna.Open fiabfr, DBCn, adOpenDynamic, adLockReadOnly
Do While Not raAna.EOF
 aktDC.obVorb = True
 Select Case raAna!Feld
  Case "Folgeerkrankung"
   aktDC.FE(raAna!FeldNr) = True
  Case "SchulEmpfohlen" ' 0 = Diabetes, 1 = Hypertonie, 2 = keine
   aktDC.SE(raAna!FeldNr) = raAna!Zeitpunkt
  Case "SchulWahrgenommen" ' Schulungen bereits vor Einschreibung wahrgenommen: 0 = Diabetes, 1 = Hypertonie, 2 = keine
   aktDC.SW(raAna!FeldNr) = raAna!Zeitpunkt
  Case "DiabSchulung" ' empfohlene Schulungen wahrgenommen, Diabetes: 0 = ja, 1 = nein, 2 = war aktuell nicht möglich
   aktDC.DS(raAna!FeldNr) = raAna!Zeitpunkt
  Case "HyperSchulung" ' empfohlene Schulungen wahrgenommen, Hypertonie: 0 = ja, 1 = nein, 2 = war aktuell nicht möglich
   aktDC.HS(raAna!FeldNr) = raAna!Zeitpunkt
  Case "LetzteSchulung" ' bei letzter Doku war keine Schulung empfohlen
   aktDC.kSE = raAna!Zeitpunkt
  Case "Glibenclamid"
   aktDC.VorM(0) = raAna!FeldNr
  Case "Metformin"
   aktDC.VorM(1) = raAna!FeldNr
  Case "Glucosidase"
   aktDC.VorM(2) = raAna!FeldNr
  Case "Sulfonylharnstoffe"
   aktDC.VorM(3) = raAna!FeldNr
  Case "Glitazone"
   aktDC.VorM(4) = raAna!FeldNr
  Case "Insulin"
   aktDC.VorM(5) = raAna!FeldNr
  Case "Insulinanaloga"
   aktDC.VorM(6) = raAna!FeldNr
  Case "HMG"
   aktDC.VorM(7) = 0
  Case "Antihypertensive", "Antihypertensiva"
   aktDC.VorM(8) = 0
   aktDC.VorM(13) = raAna!FeldNr ' kleine Unschärfe
  Case "Thrombozyten"
   aktDC.VorM(9) = 0
  Case "ACEHemmerDM2"
   aktDC.VorM(10) = 0
  Case "BetablockerDM2"
   aktDC.VorM(11) = 0
  Case "AntihypertensivDM2"
   aktDC.VorM(12) = 0
 End Select
 raAna.Move 1
Loop
raAna.Close
End If 'not ohneVorDMP
aktDC.RRsyst = 0
aktDC.RRdiast = 0
On Error GoTo fehler
'Set rMA = TabÖff("MedArten", "Medikament")
If obAdo Then
 raAna.Open "SELECT patAlter(ana.pat_id) PAlter, ana.* FROM `anamnesebogen` ana WHERE pat_id = " & pid, DBCn, adOpenKeyset, adLockReadOnly
 rsAnam.Open "SELECT * FROM `namen` WHERE pat_id = " & pid, DBCn, adOpenKeyset, adLockReadOnly
Else
' Set rEi = TabÖff("eintraege", "Auswahl")
''Set rlau = TabÖff("Labor", "WertSuch")
' Set rDT = TabÖff("Diagnosen")
' Set rKH = TabÖff("KHEinweis", "Auswahl")
' Set rsAnam = TabÖff("Anamnesebogen", "pat_id")
' Call raEi.Open("SELECT * FROM `eintraege` WHERE pat_id = " & PID, DBCn, adOpenDynamic, adLockReadOnly)
 'Call raDT.Open("SELECT * FROM diagnosen WHERE pat_id = " & PID, DBCn, adOpenDynamic, adLockReadOnly)
 'Call raKH.Open("SELECT * FROM kheinweis WHERE pat_id = " & PID, DBCn, adOpenDynamic, adLockReadOnly)
' Call raAna.Open("SELECT -dialyse AS j_dialyse, a.* FROM `anamnesebogen` a WHERE pat_id = " & pid, DBCn, adOpenDynamic, adLockReadOnly)
 Call raAna.Open("SELECT patAlter(ana.pat_id) PAlter, ana.* FROM `anamnesebogen` ana WHERE pat_id = " & pid, DBCn, adOpenDynamic, adLockReadOnly)
 Call rsAnam.Open("SELECT * FROM `namen` WHERE pat_id = " & pid, DBCn, adOpenDynamic, adLockReadOnly)
' DoCmd.Save acForm, Anmnb geht leider nicht
'If PID = 0 Then
' Dim tonRunde%
' For tonRunde = 1 To 10
'  Call Sound(WinDir + "\media\Windows XP-Standard.wav")
' Next tonRunde
' MsgBox "Achtung: DMP-Infos mit PID=0!"
' ON Error Resume Next
' PID = Forms!Anamnesebogen!Pat_id
' ON Error GoTo fehler
' If PID = 0 Then
'  For Each tbl In aktDCb.TableDefs
'   DoCmd.Close acTable, tbl.Name, acSaveYes
'  Next
'  DoCmd.OpenForm aktDCb.Containers(2).Documents(0).Name
'  DoCmd.Maximize
'  PID = Forms!Anamnesebogen!Pat_id
' End If
'End If
End If
Call rfal.Open("SELECT " & IIf(Not LVobMySQL, "top 1", "") & " * FROM `faelle` WHERE pat_id = " & pid & " AND bhfb <= " & DatFor_k(DokuDat) & " ORDER BY bhfb DESC, schgr" & IIf(LVobMySQL, " LIMIT 1", ""), DBCn, adOpenStatic, adLockReadOnly)

aktDC.Pat_id = pid
If IsNull(raAna!Nachname) Then ' so am 2.9.08 ereignet
 aktDC.Nachname = rsAnam!Nachname
 aktDC.Vorname = rsAnam!Vorname
Else
 aktDC.Nachname = raAna!Nachname
 aktDC.Vorname = raAna!Vorname
End If
Pat_id = pid ' Pat_id = modulübergreifend
aktDC.GebDat = rsAnam!GebDat
aktDC.Titel = rsAnam!Titel
aktDC.plz = rsAnam!plz
aktDC.ort = rsAnam!ort
aktDC.Lkz = rsAnam!Lkz
aktDC.PFPlz = rsAnam!PFPlz
aktDC.PFOrt = rsAnam!PFOrt
aktDC.NVors = rsAnam!NVors
aktDC.PrivatTel = rsAnam!PrivatTel
If Not IsNull(rsAnam!NVorsatz) Then aktDC.NVorsatz = rsAnam!NVorsatz
aktDC.Geschlecht = rsAnam!Geschlecht
aktDC.Anschrzus = rsAnam!Anschrzus
If aktDC.Lkz = "" Then aktDC.Lkz = "D"
aktDC.Postleitzahl = "D " & aktDC.plz & " " & aktDC.ort ' 1.1.15 Länderkennzeichen
aktDC.PrivatTel = rsAnam!PrivatTel
#If False Then ' 29.6.15, Woltmann
 Call rform.Open("SELECT " & IIf(Not LVobMySQL, "top 1", "") & " feldinh FROM `formular` WHERE pat_id = " & Pat_id & " AND Feld = 'KVKGueltig' AND zeitpunkt <= " & DatFor_k(Min(Now(), QEnd(ZQuart(Now - Verspätung)))) & " AND feldinh like '%/%'" & " ORDER BY zeitpunkt DESC" & IIf(LVobMySQL, " LIMIT 1", ""), DBCn, adOpenStatic, adLockReadOnly)
 If Not rform.BOF Then
  Postleitzahl = Left$(Postleitzahl & Space$(25), 25) & rform!FeldInh
 Else
  Postleitzahl = Left$(Postleitzahl & Space$(30), 30)
 End If
#End If

 aktDC.KVKStatus = rsAnam!KVKStatus
 
 aktDC.strasse = rsAnam!strasse
 aktDC.Hausnr = rsAnam!Hausnr
 aktDC.Versichertennummer = Trim$(rsAnam!Versichertennummer)
 If Not rfal.BOF Then
  aktDC.VKNr = rfal!VKNr
  aktDC.IK = rfal!IK
  aktDC.VschBeg = rfal!VschBeg
  aktDC.KtrAbrB = rfal!KtrAbrB
  aktDC.dmpVertret = LCase$(rfal!dmpVertret)
  aktDC.dmpArztw = LCase$(rfal!dmpArztw)
  aktDC.dmpHypos = rfal!dmpHypos
  aktDC.dmpKhsA = rfal!dmpKhsA
  aktDC.dmpDMSchulEmpf = LCase$(rfal!dmpDMSchulEmpf)
  aktDC.dmpDMSchulWahrg = LCase$(rfal!dmpDMSchulWahrg)
  aktDC.dmpHypertSchulEmpf = LCase$(rfal!dmpHypertSchulEmpf)
  aktDC.dmpHypertSchulWahrg = LCase$(rfal!dmpHypertSchulWahrg)
  aktDC.dmpKKTabakEmpf = LCase$(rfal!dmpKKTabakEmpf)
  aktDC.dmpKKErnEmpf = LCase$(rfal!dmpKKErnEmpf)
  aktDC.dmpKKkTrainEmpf = LCase$(rfal!dmpKKkTrainEmpf)
  aktDC.dmpHbA1cZiel = Replace$(rfal!dmpHbA1cZiel, ".", ",")
 End If
 If Not IsNumeric(aktDC.dmpHbA1cZiel) Then
  Dim rHb As New ADODB.Recordset
  rHb.Open "SELECT replace(dmphba1cziel,'.',',') FROM faelle WHERE pat_id = " & pid & " AND dmphba1cziel REGEXP '^-?[+]?[.,0-9]+$' ORDER BY bhfb DESC", DBCn, adOpenStatic, adLockReadOnly
  If Not rHb.BOF Then
   aktDC.dmpHbA1cZiel = rHb.Fields(0)
  End If
  Set rHb = Nothing
 End If
 If Not rfal.BOF Then
  aktDC.dmpUewFuss = LCase$(rfal!dmpUewFuss)
  aktDC.dmpEinwDM = LCase$(rfal!dmpEinwDM)
 End If
 aktDC.dmphalbj = "n" ' LCase$(rfal!dmphalbj)
 Dim rDMPh As New ADODB.Recordset
  rDMPh.Open "SELECT MAX(dmphalbj) FROM faelle f0 WHERE pat_id = " & pid & " AND bhfb = (SELECT MAX(bhfb) FROM faelle f WHERE f.pat_id = " & pid & " AND dmphalbj<>'')", DBCn, adOpenStatic, adLockReadOnly
  If Not rDMPh.BOF Then If Not IsNull(rDMPh.Fields(0)) Then aktDC.dmphalbj = rDMPh.Fields(0)
 Set rDMPh = Nothing
'Dim sqllabor$
'sqllabor$ = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, Abkü, Langtext, Wert, Einheit, Kommentar, AbsPos, AktZeit FROM `laborlangtext` INNER JOIN (laborkommentar INNER JOIN `laborneu` ON laborkommentar.KommentarVW = `laborneu`.KommentarVW) ON `laborlangtext`.LangtextVW = `laborneu`.LangtextVW WHERE pat_id = " & CStr(pid)
''Const sqllabor$ = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, Abkü, Langtext, Wert, Einheit, Kommentar, AbsPos, AktZeit FROM `laborlangtext` INNER JOIN (`laborneu`) ON `laborlangtext`.LangtextVW = `laborneu`.LangtextVW"
'lsql = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" AS NB FROM (" & sqllabor & ") AS labor WHERE wert <> """" AND not ISNULL(wert)"
'lsql = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" AS NB FROM (" & sqllabor & ") AS labor WHERE wert <> """" AND not ISNULL(wert) UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar,Normbereich AS NB " + _
 "FROM `laborxus` LEFT JOIN laborxwert ON laborxus.RefNr=laborxwert.RefNr " + _
 "WHERE pat_id = " & aktDC.Pat_id & " AND wert <> """" AND not ISNULL(wert) AND NOT EXISTS (SELECT * FROM `laborneu` WHERE pat_id = " & aktDC.Pat_id & " AND abkü = laborxwert.Abkü AND wert = laborxwert.wert AND zeitpunkt > laborxus.Eingang-3 AND zeitpunkt < laborxus.Eingang+6)"
'lsql = "SELECT * FROM labor1 WHERE pat_id = " & Pat_id & " UNION SELECT * FROM labor2 WHERE pat_id = " & Pat_id
'lsql = "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_id & " UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_id & ") i GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb"
'lsql = "SELECT lab.* FROM (SELECT @wertkl:=0 @zpkl=0 @abkü='' @patid:= " & Pat_id & " nix) nul, geslab lab"
'ralauSql = "SELECT * FROM (" & lsql & ") AS innen ORDER BY zeitpunkt DESC"

'Call raLaU.Open(ralauSql, DBCn, adOpenDynamic, adLockReadOnly)
'Do While Not raLaU.EOF
' Debug.Print raLaU!Abkü & " " & raLaU.Fields(0) & " " & raLaU.Fields(1) & " " & raLaU.Fields(2) & " " & raLaU.Fields(3) & " " & raLaU.Fields(4) & " " & raLaU.Fields(5)
' raLaU.Move 1
'Loop

'Set ralau = aktDCb.OpenRecordset("SELECT * FROM `" + QMdbAkt + "`.laborUNION WHERE pat_id = " + CStr(aktDC.pat_id) + " ORDER BY zeitpunkt DESC")
'With Forms!anamnesebogen
GesNa = GesNamFn(rsAnam) & ", *" & Format$(raAna!GebDat, "dd.mm.yyyy") & " (Pat'nr " & rsAnam!Pat_id & ")" & IIf(mitBezeich, ", DMP-Informationen für " & Format(DokuDat, "d.m.yy"), vNS)
If mitBezeich Then
 GesNa = GesNa
End If
With raAna
 If mitStr Then DmPStrS = IIf(mitBezeich, vNS, "DMP-Informationen vom " & Format$(Date, "DD.MM.YY") & " zu " & vbTab & vbTab & vbTab & vbTab & vbCrLf) & GesNa + ":"
 
 If Not IsNull(raAna!Vorgestellt) Then
  aktDC.daseit = zuJahr(DSeit(raAna, True), raAna!Vorgestellt)
  If mitStr Then TabPr "D.m. Typ " + dtyp(raAna!Diabetestyp) & IIf(aktDC.dtyp <> raAna!Diabetestyp, " (laut Anamnesebogen; " & aktDC.dtyp & " laut ICD)", vNS) & " seit:", aktDC.daseit
 End If
 
 #If vorca2008 Then
  If obPosi(![Unterzucker pM]) Or obPosi(![keto pa]) Or obPosi(![Fremde Hilfe pa]) Or obPosi(![Bewußtlos pa]) Then
   aktDC.dspsy = True
   If mitStr Then TabPr "Diagnosespez. Symptome:", "ja"
  End If
  If diI("Z34 Z33") Then
   If mitStr Then TabPr "Schwangerschaft laut ICD-Diagnosen:", "ja"
   aktDC.obSchw = True
  End If
 #End If ' vorca2008
 
 Dim VorDat As Date '(nur Attrappe)
 Call TherAuskunft(CStr(pid), 0, aktDC.insz, VorDat, aktDC.obIns, aktDC.obAnal, aktDC.obGlib, aktDC.obmetf, aktDC.obGlucI, aktDC.obSHGlin, aktDC.obGlit, aktDC.obdpp4, aktDC.obglp1, aktDC.obsglt2, aktDC.obSonstAD, aktDC.obHMG, aktDC.obAntihyp, aktDC.obACEH, aktDC.obBetabl, aktDC.obThro, aktDC.obOAK, , , , aktDC.obDiur, aktDC.obAT1)
 'if mitstr then TabPr "Diagnosespez. Medik.: ", IIf(!`DiabetesMedikament 1` <> "" AND Not ISNULL(!`DiabetesMedikament 1`), "ja", "nein")
 If aktDC.dtyp = "1" And Not aktDC.obIns And Not aktDC.obAnal Then ' Medikamentenplan fehlt oder keine Dosierung angegeben
  Dim rez As New ADODB.Recordset, rezSQL$
  rezSQL = "SELECT * FROM (SELECT *,left(medikament,instr(medikament,"" "")-1) AS medanf FROM `rezepteintraege` ) AS rez LEFT JOIN `medarten` ON rez.medanf = `medarten`.medikament WHERE rez.pat_id = 150 AND (ins Or anal) ORDER BY zeitpunkt DESC;"
'  If Not lies.obMySQL Then rezSQL = Replace$(rezSQL, "collate latin1_german2_ci ", vNS)
  rez.Open rezSQL, DBCn, adOpenStatic, adLockReadOnly
  If Not rez.BOF Then
   aktDC.obIns = rez!InS
   aktDC.obAnal = rez!anal
  End If
  rez.Close
 End If
 
 #If vorca2008 Then
  aktDC.dspmed = (aktDC.obIns Or aktDC.obAnal Or aktDC.obGlib = adja Or aktDC.obmetf = adja Or aktDC.obGlucI = adja Or aktDC.obSHGlin = adja Or aktDC.obGlit = adja Or aktDC.obdpp4 = adja Or aktDC.obglp1 = adja Or aktDC.obsglt2 = adja Or aktDC.obSonstAD = adja)
' aktDC.dspmed = (aktDC.VorM(0) Or aktDC.VorM(1) Or aktDC.VorM(2) Or aktDC.VorM(3) Or aktDC.VorM(4) Or aktDC.VorM(5) Or aktDC.VorM(6))
  If mitStr Then TabPr "Diagnosespez. Medik.:", IIf(aktDC.dspmed, "ja", "nein")
 #End If ' vorca2008
 
' 2.Reiter
 Dim bbk$, spät$, ereig$
 bbk = vNS
 spät = vNS
 ereig = vNS
 'If instrb(!Diagnosen, "yperton") <> 0 Or instrb(!Diagnosen, "ochdru") <> 0 Then bbk = bbk + " Hypertonus,"
 'If diT("yperton") Or diT("ochdru") Then bbk = bbk + " Hypertonus,"
 If diI("I10") Then
  aktDC.FEn(1) = True
  bbk = bbk + " Art.Hypertonie" & IIf(aktDC.FE(1) And Not ab315, "( bek)", "") & ","
 End If
 If diI("I50") Then ' Herzinsuffizienz
  aktDC.FEn(17) = True
  bbk = bbk + " Chron.Herzinsuff." & IIf(aktDC.FE(17) And Not ab315, "( bek)", "") & ""
 End If
 If diT("ettstoffw yperchol") Or diI("E78") Then
  aktDC.FEn(2) = True
  bbk = bbk + " Fettstoffw'strg." & IIf(aktDC.FE(2) And Not ab315, "( bek)", "") & ","
 End If
 If diI("J45") Then ' Asthma
  aktDC.FEn(16) = True
  bbk = bbk + " Asthma bronch." & IIf(aktDC.FE(16) And Not ab315, "( bek)", "") & ""
 End If
 If diI("J42 J44") Then ' COPD
  aktDC.FEn(15) = True
  bbk = bbk + " COPD" & IIf(aktDC.FE(15) And Not ab315, "( bek)", "") & ""
 End If
 If diI("I7") Then ' pAVK usw
  aktDC.FEn(6) = True
  bbk = bbk + " pAVK" & IIf(aktDC.FE(6) And Not ab315, "( bek)", "") & ","
 End If
 If diI("I20 I21 I22 I23 I24 I25") Then ' I20 A.p., I21 akuter MI, I22 rez.MI, I23 Kompl.n.MI, I24 sonst.ischäm.Hkt, I25 Atherosk.Hkt, alter MI
  aktDC.FEn(3) = True
  bbk = bbk + " Koronare Herzerkrankung" & IIf(aktDC.FE(3) And Not ab315, "( bek)", "") & ","
 End If
  
 If Not ab315 Then If LenB(bbk) <> 0 Then bbk = bbk & vbCrLf & Space$(BreiteSp1 - 1)
 Dim ICD$, diagdat As Date
 If obPosi(!Herzinfarkt) Or diI("I21 I22", , , True, , aktDC.FEi(4), aktDC.FEd(4)) Then
  aktDC.FEn(4) = True
  If Not ab315 Then bbk = bbk + " Herzinfarkt" & IIf(aktDC.FE(4) And Not ab315, "( bek)", "") & ","
 End If
 If diT("poplex chlaganf erebral", , , True, , aktDC.FEi(5), aktDC.FEd(5)) Or _
    diI("I63 I64", , , True, , aktDC.FEi(5), aktDC.FEd(5)) Then
  aktDC.FEn(5) = True
  If Not ab315 Then bbk = bbk + " Schlaganfall" & IIf(aktDC.FE(5) And Not ab315, "( bek)", "") & ","
 End If
 If diI("N08.3 N18 N19") And Not obkNeph Then
  aktDC.FEn(7) = True
  If ab315 Then
   spät = "Nephropathie"
  Else
   bbk = bbk + " Nephropathie" & IIf(aktDC.FE(7) And Not ab315, "( bek)", "") & ","
  End If
 End If
  
' If obPosi(!j_Dialyse) OR diI("Z49.1") Then
 If obPosi(!Dialyse <> 0) Or diI("Z49", , , , , aktDC.FEi(8), aktDC.FEd(8)) Then
  aktDC.FEn(8) = True
  If ab315 Then
   spät = IIf(LenB(spät) = 0, "Nephropathie:", ":") & " Nierenersatztherapie"
  Else
   bbk = bbk + " Nierenersatztherapie" & IIf(aktDC.FE(8) And Not ab315, "( bek)", "") & ","
  End If
 End If
 
  aktDC.obG632 = diI("G63.2")
  If aktDC.obG632 Or diI("G59.0 G99.0") Or diT("europath") Then
   aktDC.FEn(11) = True
   If ab315 Then
    spät = spät & IIf(LenB(spät) = 0, "", ", ") & "Neuropathie"
   End If
  End If
  
  If diT("Retinopath") Or diI("H36") Then
   aktDC.FEn(9) = True
   If ab315 Then
    spät = spät & IIf(LenB(spät) = 0, "", ", ") & "Retinopathie"
   Else
    bbk = bbk + " diab. Retinopathie" & IIf(aktDC.FE(9) And Not ab315, "( bek)", "") & ","
   End If
  End If
  If InStrB(LCase(!Diagnosen), "blind") <> 0 Or diI("H54 S05", , , True, , aktDC.FEi(10), aktDC.FEd(10)) Then
   aktDC.FEn(10) = True
   If ab315 Then
    spät = spät & IIf(aktDC.FEn(9), ", ", IIf(LenB(spät) = 0, "", ", ") & "Retinopathie") & " (Blindheit)"
   Else
    bbk = bbk + " Blindheit" & IIf(aktDC.FE(10) And Not ab315, "( bek)", "") & ","
   End If
  End If
      
  If Not ab315 Then
   If LenB(bbk) <> 0 Then bbk = bbk & vbCrLf & Space$(BreiteSp1 - 1)
   If aktDC.FEn(11) Then
    bbk = bbk + " diab. Neuropathie" & IIf(aktDC.FE(11) And Not ab315, "( bek)", "") & ","
   End If
  End If
  
  If diI("L89 M14.6") Then
   aktDC.FEn(12) = True
   If Not ab315 Then bbk = bbk + " Diabetisches Fußsyndrom" & IIf(aktDC.FE(12) And Not ab315, "( bek)", "") & ","
  End If
  If obPosi(!Amputation) Or diT("mputat", , , True, , aktDC.FEi(13), aktDC.FEd(13)) Then
   aktDC.FEn(13) = True
   If Not ab315 Then bbk = bbk + " Amputation" & IIf(aktDC.FE(13) And Not ab315, "( bek)", "") & ","
  End If

  If diI("M36.8 M14.2") Or InStrB(!Diagnosen, "aculop") <> 0 _
  Or diI("K77.8 K71.7 K71.6 K76.9 K76.2") Or diT("atara") > 0 Or diT("arnwegsinf") > 0 Then
   aktDC.FEn(14) = True
   If Not ab315 Then bbk = bbk + " Sonstige" & IIf(aktDC.FE(14) And Not ab315, "( bek)", "") & ""
  End If
  
 Dim j%
 aktDC.FEn(0) = -1
 For j = 1 To UBound(aktDC.FEn)
  If aktDC.FEn(j) And Not aktDC.FE(j) Then
   aktDC.FEn(0) = 0
   Exit For
  End If
 Next j
 If bbk = "" Then bbk = " keine"
 If Right$(bbk, 1) = "," Then bbk = Left(bbk, Len(bbk) - 1)
 Do While Asc(Right$(bbk, 1)) = 9 Or Asc(Right$(bbk, 1)) = 10 Or Asc(Right$(bbk, 1)) = 11 Or Asc(Right$(bbk, 1)) = 13
  bbk = Left(bbk, Len(bbk) - 1)
  If Len(bbk) = 0 Then Exit Do
 Loop
 bbk = LTrim$(bbk)
 If mitStr Then TabPr "•  Bek. Begleit-/Folgeerk.:", bbk
' Dim TabakExpr$
' If ISNULL(!Tabak) Or !Tabak = "" Then
'  TabakExpr = vns
' Else
'  If instrb("jJ", trim$(!Tabak)) > 0 Then
'   TabakExpr = "ja"
'  ElseIf instrb("-nN0", trim$(!Tabak)) > 0 Then
'   TabakExpr = "nein"
'  Else
'   TabakExpr = CStr(!Tabak)
'  End If
' End If
 aktDC.obRauch = diI("F17.1")
 aktDC.Tabak = (WieTabak(!Pat_id) = aktuell)
 If mitStr Then TabPr "Raucher:", IIf(aktDC.Tabak, "ja", "nein")
 If Not IsNull(!Größe) Then aktDC.kgr = Round(Replace$(IIf(!Größe < 10, !Größe * 100, !Größe), ".", ","), 0)
 If Not IsNull(!Größe) Then If mitStr Then TabPr "Körpergröße:", CStr(!Größe) + IIf(!Größe < 10, " m", " cm")
 
 raEi.Open "SELECT * FROM `eintraege` WHERE pat_id = " & pid & " AND art = ""Gewicht"" ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
 If Not raEi.EOF Then
  Dim gewi$
  gewi = Trim$(Replace$(Replace$(LCase(raEi!Inhalt), "kg", ""), ".", ","))
  If InStrB(gewi, "-") <> 0 Then gewi = Left(gewi, InStr(gewi, "-") - 1)
   Dim pos&
  Do
   If Not IsNumeric(gewi) Then
    pos = InStr(gewi, " ")
    If pos > 0 Then
     If IsNumeric(Left(gewi, pos - 1)) Then
      gewi = Left(gewi, pos - 1)
     Else 'If IsNumeric(mid$(gewi, pos + 1)) Then
      gewi = Mid$(gewi, pos + 1)
     End If
    Else
     For j = 33 To 255
      If (j < 48 Or j > 57) And j <> 44 Then
       gewi = Replace$(gewi, Chr$(j), "")
      End If
     Next j
    End If
   End If
   If InStr(gewi, ",") > 0 Then
    If InStr(InStr(gewi, ",") + 1, gewi, ",") > 0 Then
     gewi = Left(gewi, InStr(InStr(gewi, ",") + 1, gewi, ",") - 1)
    End If
   End If
   gewi = Replace$(gewi, ",,", ",")
   If gewi = "" Or gewi = "," Then
    gewi = "0"
    Exit Do
   End If
   If IsNumeric(gewi) Then Exit Do
  Loop
  If Not IsNumeric(gewi) Then
   aktDC.gewi = 0
  Else
   aktDC.gewi = Round(gewi, 0)
  End If
  If mitStr Then TabPr "Körpergewicht:", raEi!Inhalt + IIf(InStrB(raEi!Inhalt, "kg") <> 0, vNS, " kg") & " (" & CStr(DateValue(raEi!Zeitpunkt)) & ")"
 Else
  If Not IsNull(!Gewicht) Then
   aktDC.gewi = Round(Replace$(!Gewicht, ".", ","), 0) ' ist double
   If mitStr Then TabPr "Körpergewicht:", CStr(!Gewicht) + " kg"
  End If
 End If ' Not raEi.EOF Then else
    
 If aktDC.kgr <> 0 Then
  aktDC.bmi = aktDC.gewi / (aktDC.kgr * 1# * aktDC.kgr)
  If aktDC.bmi > 0 Then
   Do While aktDC.bmi < 10
    aktDC.bmi = aktDC.bmi * 10
   Loop
   Do While aktDC.bmi > 100
    aktDC.bmi = aktDC.bmi * 0.1
   Loop
  End If
 End If ' aktDC.kgr <> 0 Then
 
 aktDC.PrRR = GetPrRR(raAna, aktDC.RRsyst, aktDC.RRdiast, obdiastkorr:=True)
 If aktDC.RRdiast < 30 Then aktDC.RRdiast = Min(70, aktDC.RRsyst) ' 31.10.20
 If aktDC.dtyp = "2" Then
  If aktDC.PrRR <> "" Then
   If mitStr Then TabPr "Blutdruck:", aktDC.PrRR
  End If
 End If
 
If True Then ' lwZahl
 aktDC.bekHb = 0
' ralau.Seek "=", CStr(PID), "HBA1C"
 Dim DMPHbA1c$, DMPCrea$, DMPeGFR$, DMPUAlb$
' Set raLau = Nothing
' raLau.Open "SELECT * FROM (" & lsql & ") AS innen WHERE abkü = ""HBA1C"" ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
' Set raLau = LabEPat(HbA1c, PID)
 alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
 Labs = LabPat(LA_HbA1c, pid)
 Debug.Print Labs.WertSg
' raLaU.MoveFirst
' raLaU.Find "Abkü = ""HBA1C"""
'  DMPHbA1c = raLau!Wert & "% (" & Format$(raLau!Zp, "dd/mm/yy") & ")" & IIf(Not ab315, ", oberer Normwert: 6,2%", "")
 
' Set raLaum = LabEPat(HbA1cmm, PID)
 alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
 LabSm = LabPat(LA_HbA1cmm, pid)

 Dim Quelle% '0= raLau, 1=raLaum, 2= Formular
' If Not raLau.BOF AND raLaum.BOF Then
' ElseIf raLau.BOF AND Not raLaum.BOF Then
 If Labs.Abkü = "" And LabSm.Abkü <> "" Then
  Quelle = 1
' ElseIf raLau.BOF AND raLaum.BOF Then
 ElseIf Labs.Abkü = "" And LabSm.Abkü = "" Then
  Quelle = 2
' ElseIf raLau!zp >= raLaum!zp Then
' ElseIf raLau!zp < raLaum!zp Then
 ElseIf Labs.Zp <= LabSm.Zp Then
  Quelle = 1
 Else
  Quelle = 0
 End If
 If Quelle = 1 Then
'  DMPHbA1c = Round((raLaum!wert * 0.0915 + 2.15), 1) & " % / " & raLaum!wert & " mmol/mol (" & Format$(raLaum!zp, "dd/mm/yy") & ")"
  aktDC.bekHb = Round((LabSm.WertSg * 0.0915 + 2.15), 1)
  DMPHbA1c = aktDC.bekHb & " % / " & LabSm.WertSg & " mmol/mol (" & Format$(LabSm.Zp, "dd/mm/yy") & ")"
 ElseIf Quelle = 0 Then
'  If raLau!wert <> "" Then
  If Labs.WertSg <> "" Then
   If InStrB(Labs.WertSg, " ") <> 0 Then
    pos = InStr(Labs.WertSg, " ")
    Dim zwStri$
    zwStri = Replace$(Left$(Labs.WertSg, pos - 1), ".", ",")
    If IsNumeric(zwStri) Then
     aktDC.bekHb = Round(CDbl(zwStri), 1)
    End If
   Else
    Dim rLauWert$
    rLauWert = Replace$(Labs.WertSg, "?", vNS)
    pos = InStr(rLauWert, ">") ' Pat_id 744
    If pos > 0 Then rLauWert = Mid(rLauWert, pos + 1)
    aktDC.bekHb = Round(CDbl(Replace$(rLauWert, ".", ",")), 1)
   End If
  End If
'  DMPHbA1c = raLau!wert & " % / " & Round((aktDC.bekHb - 2.15) * 10.929) & " mmol/mol (" & Format$(raLau!zp, "dd/mm/yy") & ")"
'  DMPHbA1c = LabS.wertsg & " % / " & Round((aktDC.bekHb - 2.15) * 10.929) & " mmol/mol (" & Format$(LabS.zp, "dd/mm/yy") & ")"
  Dim LabsWert!
  LabsWert = IIf(Labs.WertSg = vNS, 0, aktDC.bekHb)
  DMPHbA1c = LabsWert & " % / " & Round((LabsWert - 2.15) * 10.929) & " mmol/mol (" & Format$(Labs.Zp, "dd/mm/yy") & ")"
 ElseIf Quelle = 2 And tfeld(![letztes HbA1c]) <> "" Then
  Dim testHbA1c$
  testHbA1c = Replace$(Replace$(![letztes HbA1c], ".", ","), "%", "")
  If IsNumeric(testHbA1c) Then
   aktDC.bekHb = Round(Replace$(testHbA1c, ".", ","), 1)
  End If
  DMPHbA1c = ![letztes HbA1c] + " " + IIf(LenB(tfeld(![gemessen am])) = 0, "(auswärtig)", "( " + tfeld(![gemessen am]) + ")")
 End If
 
 
 ' DMPHbA1c = DMPHbA1c & " / " & raLau!Wert & " mmol/l (" & Format$(raLau!Zp, "dd/mm/yy") & ")"
 
' Set raLau = Nothing
' Set raLau = LabEPat(Krea, PID)
 alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
 Labs = LabPat(la_krea, pid)
' If Not raLau.EOF Then
 If Labs.Abkü <> "" Then
'  aktDC.Crea = Round(Replace$(Replace$(IIf(Right$(raLau!Wert, 4) = "Jaffe", Left$(raLau!Wert, Len(raLau!Wert) - 4), raLau!Wert), ".", ","), ";", ","), 1)
  Dim CreaR$
  CreaR = Labs.WertSg
  If Right$(CreaR, 5) = "Jaffe" Then
   CreaR = Left$(CreaR, Len(CreaR) - 5)
  End If
  CreaR = Replace$(CreaR, ".", ",")
  CreaR = Replace$(CreaR, ";", ",")
  CreaR = Replace$(CreaR, "(?)", vNS)
  CreaR = Replace$(CreaR, "?", vNS)
  CreaR = Replace$(CreaR, "(S)", "")
  If CreaR = "" Then CreaR = "0"
  If InStrB(CreaR, " ") <> 0 Then CreaR = Left$(CreaR, InStr(CreaR, " ") - 1)
  If CreaR = "o,B," Then CreaR = "1.0"
  
  Dim pospfeil%
  pospfeil = InStr(CreaR, ">")
  If pospfeil > 0 Then CreaR = Mid(CreaR, pospfeil + 1)
  aktDC.Crea = Replace$(CreaR, ",,", ",")
  aktDC.Crea = Round(aktDC.Crea, 1)
  DMPCrea = Labs.WertSg & " mg/dl (" & Format$(Labs.Zp, "dd/mm/yy") & ")" ', oberer Normwert: 1,3 mg/dl"
 End If
 If ab315 Then
 ' das folgende umgeändert am 15.4.20
''  Set raLau = Nothing
''  Set raLau = LabEPat(eGFR, PID)
'  alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
'  Labs = LabPat(LA_eGFR, PID)
''  If Not raLau.EOF Then
'  If Labs.Abkü <> "" Then
'   Dim eGFRs$
'   eGFRs = Labs.WertSg
'   eGFRs = Replace$(eGFRs, ".", ",")
'   eGFRs = Replace$(eGFRs, ";", ",")
'   eGFRs = Replace$(eGFRs, "(?)", ",")
'   eGFRs = Replace$(eGFRs, "?", ",")
'   eGFRs = Replace$(eGFRs, "(S)", ",")
'   If eGFRs = "" Then eGFRs = "0"
'   If InStrB(eGFRs, " ") <> 0 Then eGFRs = Left$(eGFRs, InStr(eGFRs, " ") - 1)
'   If eGFRs = "o,B," Or eGFRs = ">60" Then eGFRs = "80"
'   If eGFRs = ">" Then aktDC.eGFR = 80 Else aktDC.eGFR = CLng(eGFRs)
'
'   aktDC.eGFR = Round(aktDC.eGFR, 0)
'   If aktDC.eGFR = 0 Then
'    If aktDC.Crea = 0 Then
'     aktDC.eGFR = 0
'    Else
'     aktDC.eGFR = Round((140 - raAna!palter) * aktDC.gewi * IIf(aktDC.gewi < 3, 100, 1) * IIf(aktDC.Geschlecht = "w", 0.85, 1) / aktDC.Crea / 72, 0)
'    End If
'   End If
'   DMPeGFR = IIf(aktDC.eGFR, aktDC.eGFR & " ml/min (" + Format$(Labs.Zp, "dd/mm/yy") + ")", "nicht vorliegend") ', oberer Normwert: 1,3 mg/dl"
'  End If
  Dim lG As labtyp
  lG = letztGFR(pid, raAna!Palter, IIf(aktDC.Geschlecht = "w", -1, 0))
  aktDC.eGFR = IIf(lG.WertSg = "", 500, lG.WertSg)
  DMPeGFR = IIf(aktDC.eGFR, aktDC.eGFR & " " & lG.Einheit & " (" + Format$(lG.Zp, "dd/mm/yy") + ")", "nicht vorliegend") ', oberer Normwert: 1,3 mg/dl"
 End If ' ab315 Then
 
' If aktDC.dtyp = "1" Then ' 5.10.08: jetzt auch Typ 2
'  Set raLau = Nothing
'  raLau.Open "SELECT * FROM (" & lsql & ") AS innen WHERE abkü in ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND (abkü <> 'ALBU' Or wert like '%<%') ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
'  sql = "SELECT zeitpunkt zp, abkü, einheit, wert FROM `laborneu` WHERE abkü in ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND pat_id = " & Pat_id & " UNION SELECT u.eingang zp, w.abkü abkü, w.einheit einheit, w.wert wert FROM `laborxus` u LEFT JOIN laborxwert w ON u.refnr = w.refnr WHERE abkü in ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND pat_id = " & Pat_id & " ORDER BY pat_id,zp DESC"
'  Set raLau = LabEPat(AlbCre, PID)
  alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
  Labs = LabPat(LA_AlbCre, pid)
  ' ALBCRE, ALBKRE, ALBQ, ALBUM, ALBUP
' raLaU.Find "Abkü in (""CREAT"",""KRE02"", ""KREA"", ""KREA02"", ""KRES"")"
'  If raLau.EOF Then
  If Labs.Abkü = "" Then
   Dim rEintr As New ADODB.Recordset
   rEintr.Open "SELECT * FROM `eintraege` WHERE art = ""urin"" AND inhalt like ""%micral%"" AND Pat_ID = " & aktDC.Pat_id & " ORDER BY zeitpunkt DESC", DBCn, adOpenStatic, adLockReadOnly
   If Not rEintr.EOF Then
    DMPUAlb = Mid$(rEintr!Inhalt, InStr(LCase(rEintr!Inhalt), "micral")) & " (" & Format$(rEintr!Zeitpunkt, "dd/mm/yy") & ")"
   End If
  Else
   If InStrB(Labs.WertSg, "<") <> 0 Then
    aktDC.mau = unauff
   Else
     Dim raLauW$
     raLauW = Replace$(Replace$(Replace$(IIf(Labs.WertSg = vNS, "0", Labs.WertSg), "ca.", ""), ".", ","), " (Urin)", vNS)
     If IsNumeric(raLauW) Then
      If Round(raLauW, 1) > 20 Then
       aktDC.mau = auff
      Else
       aktDC.mau = unauff
      End If
     Else
      aktDC.mau = unauff
     End If
   End If
   Dim werts$
   If IsNumeric(Labs.WertSg) Then werts = CStr(Labs.WertSg) Else werts = Labs.WertSg
   DMPUAlb = werts + " mg/dl (" + Format$(Labs.Zp, "dd/mm/yy") + "), oberer Normwert: 29 mg/gCrea"
   If IsNumeric(Labs.WertSg) Then
    If Labs.WertSg >= 30 Then
     aktDC.mau = auff
    Else
     aktDC.mau = unauff
    End If
   End If
'  End If
 End If
 If aktDC.dtyp = "2" Then
  If mitStr Then TabPr IIf(ab317, "•  ", "") & "HbA1c:", DMPHbA1c
  If ab315 Then
   If mitStr Then TabPr "eGFR:", DMPeGFR
  Else
   If mitStr Then TabPr "Serum-Kreatinin:", DMPCrea
  End If
  If mitStr Then TabPr "Urin-Albumin:", DMPUAlb ' 5.10.08 Kommentar entfernt
 End If
 'If obPosi(!Beinbefund) Or obPosi(!Ulcera) OR diI("M14") Then
 End If 'lwzahl
 
 On Error Resume Next
 ReDim rNa(0)
 On Error GoTo fehler
 rNa(0).Pat_id = pid
 Call usdmAlt(True)
 Dim trp As Boolean, tlp As Boolean, drp As Boolean, dlp As Boolean
 Dim Mfrep%, Mflip%
 Dim KWrep%, KWlip%
 Dim VibrIKrep%, VibrIKlip%
 Dim VibrGZrep%, VibrGZlip%
 Dim summe%
 Dim Nausg$, LUSDat$
 Dim MFBegr$, KWBegr$, VibBegr$
 Dim PulsStatus$
 aktDC.Puls = PStatNeu(pid, PulsStatus)
 If mitStr And Not ab317 Then TabPr "Pulsstatus: ", PulsStatus
 Dim SensText$
 aktDC.sens = sensib(pid, SensText)
 If mitStr And Not ab317 Then TabPr "Sensibilität: ", SensText
   
 If mitStr And ab315 Then
  If LenB(spät) = 0 Then spät = "keine"
  TabPr "Spätfolgen: ", spät
 End If ' mitStr AND ab315 Then


#If vorca2008 Then
 Dim abiind&

' aktDC.Puls = unauff
' For abiind = 2 To 5
'  If Len(AbI(abiind)) > 0 Then
'   If Left$(AbI(abiind), 1) = "{" Then
'    If Right$(AbI(abiind), 1) = "}" Then
'     AbI(abiind) = ""
'    End If
'   End If
'  End If
' Next abiind
' If UBound(AbI) = 0 Or (AbI(3) = "" AND AbI(2) = "" AND AbI(5) = "" AND AbI(4) = "") Then
'  PulsParse ![Puls Atp], trp, tlp
'  PulsParse ![Puls Adp], drp, dlp
'  Nausg = " (" & tfeld(![Puls Atp]) & "," & tfeld(![Puls Adp]) & ")"
'  LUSDat = !Vorgestellt
'  If tfeld(![Puls Atp]) = "" AND tfeld(![Puls Adp]) = "" Then aktDC.Puls = ndok
' Else
'  PulsParse AbI(3) & "/" & AbI(2), trp, tlp
'  PulsParse AbI(5) & "/" & AbI(4), drp, dlp
'  Nausg = " (" & AbI(3) & "/" & AbI(2) & ", " & AbI(5) & "/" & AbI(4) & ")"
'  LUSDat = Format$(AbIDate, "d.m.yy")
' End If
 
 For abiind = 10 To 17
  If Len(AbI(abiind)) > 2 Then
   If Left$(AbI(abiind), 1) = "{" Then
    If Left$(Right$(AbI(abiind), 3), 2) = "}/" Then
     AbI(abiind) = ""
    End If
   End If
  End If
 Next abiind
 If UBound(AbI) = 0 Or (AbI(17) = "" And AbI(16) = "") Then
  MFBegr = tfeld(!Monofilamenttest)
 Else
  MFBegr = AbI(17) & "," & AbI(16)
 End If
 obMonPath MFBegr, Mfrep, Mflip
 If UBound(AbI) = 0 Or (AbI(15) = "" And AbI(14) = "") Then
  KWBegr = tfeld(![Kalt-Warm])
 Else
  KWBegr = AbI(15) & "," & AbI(14)
 End If
 obKWPath KWBegr, KWrep, KWlip
 If UBound(AbI) = 0 Or (AbI(13) = "" And AbI(12) = "" And AbI(11) = "" And AbI(10) = "") Then
  obVibPath ![Vibration IK], VibrIKrep, VibrIKlip, -1
  obVibPath ![Vibration Großzehe], VibrGZrep, VibrGZlip, -1
  VibBegr = ![Vibration IK] & "; " & ![Vibration Großzehe]
 Else
  obVibPath AbI(13) & "," & AbI(12), VibrIKrep, VibrIKlip
  obVibPath AbI(11) & "," & AbI(10), VibrGZrep, VibrGZlip
  VibBegr = AbI(13) & "," & AbI(12) & ";" & AbI(11) & "," & AbI(10)
 End If
 LUSDat = LUSDat & "):"
#End If ' vorca2008

#If vorca2008 Then
' If (trp AND drp) Or (tlp AND dlp) Or (trp AND dlp) Or (tlp AND drp) Then
'   aktDC.Puls = auff
' End If
' If aktDC.Puls = auff Then
'  If mitStr Then TabPr "Pulsstatus:", "auffällig (" & LUSDat & Nausg
' ElseIf aktDC.Puls = ndok Then
'  If mitStr Then TabPr "Pulsstatus:", "nicht dokumentiert (" & LUSDat & Nausg
' Else
'  If mitStr Then TabPr "Pulsstatus:", "unauffällig (" & LUSDat & Nausg
' End If

 summe = Mfrep + Mflip + KWrep + KWlip + VibrIKrep + VibrIKlip + VibrGZrep + VibrGZlip
 Nausg = "auffällig (" & LUSDat
 If summe > 1 Then ' Or aktDC.obG632 Then
  If Mfrep > 0 Or Mflip > 0 Then Nausg = Nausg & " Monofil.:" & MFBegr & ","
  If KWrep > 0 Or KWlip > 0 Then Nausg = Nausg & " Kalt/W.:" & KWBegr & ","
  If VibrGZrep > 0 Or VibrGZlip > 0 Then Nausg = Nausg & " Vibr." & VibBegr & ","
  aktDC.sens = auff
  If mitStr Then TabPr "Sensibilitätsprüfung:", Nausg
 ElseIf aktDC.obG632 Then
  aktDC.sens = pathdok
  If aktDC.obG632 Then Nausg = Nausg & " Diagn. G63.2"
  If mitStr Then TabPr "Sensibilitätsprüfung:", Nausg
 ElseIf tfeld(!Monofilamenttest) = "" And (tfeld(![Kalt-Warm]) = "" Or tfeld(![Kalt-Warm]) = "/5 | /5" Or tfeld(![Kalt-Warm]) = "~/5 | ~/5") And tfeld(![Vibration IK]) = "" And tfeld(![Vibration Großzehe]) = "" Then
  aktDC.sens = ndok
  If mitStr Then TabPr "Sensibilitätsprüfung:", "nicht dokumentiert"
 Else
  aktDC.sens = unauff
  If mitStr Then TabPr "Sensibilitätsprüfung:", "un" & Left$(Nausg, Len(Nausg) - 2) & ")"
 End If
#End If ' vorca2008

' inj: 0 = nicht untersucht, 1 = unauffällig, 2 = auffällig
    Dim injf$
    If aktDC.obIns = 0 And aktDC.obAnal = 0 Then
     aktDC.inj = "0"
    Else
     Dim rusdm As New ADODB.Recordset
     rusdm.Open "SELECT spritzst FROM usdm WHERE pat_id = " & pid & " AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM usdm WHERE pat_id = " & pid & " AND zeitpunkt < " & Format(DokuDat + 1, "yyyymmdd") & ")", DBCn, adOpenStatic, adLockReadOnly
     If rusdm.BOF Then
      aktDC.inj = "0"
     Else
      injf = rusdm!Spritzst
      If InStr(injf, "-|-") < 4 Or InStr(injf, "n|n") = 1 Then
       aktDC.inj = "1"
      Else
       Select Case injf
        Case "normal", "-|-", "n|n", "|", "o.B.|o.B.", "lt. Pat. o.B.|lt. Pat. o.B", "", "keine|keine", "o.B.|o.B", _
             "lt. Pat. o.B.|lt. Pat. o.B.", "nein|nein", "~|", "o.B|o.B.", "--|-", "o.B. lt. Pat.|o.B. lt. Pat", _
             "- li: -|- li: -", "o. B.|o. B.", "o.B|o.B", "lt. Pat. o.B.|-", "-|", "o.B. lt. Pat.|o.B. lt. Pat.", _
             "Lt. Pat. Bauchbereich o.B.|-", "o.B. li: o.B.|o.B. li: o.B.", "oB|oB", "|~", "~-|-", "|-", _
             "lt.Pat. o.B.|lt. Pat. o.B", "n li: n|n li: n", "o.B. lt. Pat.|-", "lt. Pat.o.B.|lt. Pat. o.B", _
             "-|.-", ".-|-", "lt. Pat. o.B.|lt.Pat. o.B", "B|o.B.", "Os|n", "keine li: keine|keine li: keine", _
             ".|.", "0.B.|o.B.", "|n", "-.|-", "--|--", "nicht eindeutig|nicht eindeutig", "e: -|-", "o.B.|B", _
             "o.B.|", "-|.", "Lt. Pat. o.B.|o.B.", "n|", "-li: -|-li: -", "lt. Pat. o.B .|lt. Pat. o.B", _
             "lt. Pat.o.B.|lt. Pat. o.B.", "lt. Pat. o:b.|lt. Pat. o.b", "|Fußbefund (Inspektion): re: normal,", _
             "o.B.|o.B:", "lt. Pat.o.B.|lt. Pat.o.B", "nicht wesentlich|nicht wesentlich", "n|~n", "o.B.,|o.B.,", _
             "-|o.B.", "n -|-", ": n|n", "-,li: -|-,li: -", "o. B.|o. B", "L.t Pat. Bauch o.B.|-", "-|--", _
             "o.B. li: o.B|o.B. li: o.B", "kein e|keine", "Bauch o.B|Bauch o.B.", "Lt. Pat. Bauch o.B.|-", _
             "o.B.li: o.B|o.B.li: o.B", "-, lki: -|-, lki: -", "n|-", "Lt. Pat. alles o.B.|-", "lt. Pat. o.B.|lt. Pat. o:B.", _
             "keine lt. Pat.|keine lt. Pat.", ".|", "o.B li: o.B|o.B li: o.B", "lt. Pat. o.B.|o.B", "nein bds.|", _
             "lt. Pat o.B.|o.B.", "~o.B.|o.B.", "lt Pat. o.B.|lt. Pat. o.B", "nei|nein", "o.B|o.,B.", _
             "- li:-|- li:-", "o.B. bds|o.B. bds", "o.B.|o. B.", "o:B.|o.B.", "o.B:|o.B:", "~o.B. lt. Pat|-", _
             "-|~", "oB. lt. Pat.|", "-  li.: -|-  li.: -", "o.B.|.o.B", "keine|keiner", "_|-", "o.B . lt. Pat.|o.B. lt. Pat", _
             ": -|-", "lt. Pat. Bauch o.B .|o.B.", "o. B.|o.B.", "lt. Pat.o.B|lt. Pat.o.B", "weich|weich", _
             "lt. Pat. o.B.|~", "o.B li: o.B.|o.B li: o.B.", "o.B:|o.B.", "keinwe|keine", "Lt. Pat. o.B:|Lt. Pat. o.B.", _
             "-  li: -|-  li: -", "keine: li: keine|keine: li: keine", "o.B:|o.B", "o.|o.B", "-li:-|-li:-", _
             "o.B,.|o.B.", "keinen|keinen", "oB. lt. Pat.|o.B. lt. Pat", "lt. Pat. o.B.|lt. Pat. o .B", "lt.Pat.o.B.|lt. Pat. o.B"
         aktDC.inj = "1"
        Case "?|?", "n.u.|n.u.", "entfällt|entfällt", "entf|entf"
         aktDC.inj = "0"
        Case Else
         aktDC.inj = "2"
       End Select
      End If
     End If
    End If
  If mitStr And ab315 And ab317 Then
   TabPr "Injektionsstellen:", IIf(aktDC.inj = 0, "nicht untersucht", IIf(aktDC.inj = 1, "unauffällig", "auffällig"))
  End If
  
' 29.8.17
sql = "SELECT IF(Mid(di.icd, 5, 1)>=2,1,if(mid(di.icd,5,1)<=1,0,if(u.ulcera like '%obfl%',0,if(u.ulcera like '%tief%',1,if(e.art is null,2,0))))) ulcus " & _
"FROM aktfvs f " & _
"LEFT JOIN eintraege e ON f.pat_id = e.pat_id AND e.art='ulcus' AND e.zeitpunkt BETWEEN qanf() AND qend() " & _
"LEFT JOIN fuss u ON f.pat_id = u.pat_id AND u.ulcera in ('obfl','tief') AND u.zeitpunkt BETWEEN qanf() AND qend() " & _
"LEFT JOIN diagnosen di ON f.pat_id = di.pat_id AND  di.gicd RLIKE '^L89\.[12345]' AND obdauer = 0 AND date(di.diagdatum) BETWEEN qanf() AND qend() " & _
"WHERE (Not ISNULL(e.Art) Or Not ISNULL(u.Ulcera) Or Not ISNULL(di.ICD)) " & _
" AND f.pat_id = " & pid & _
" GROUP BY f.pat_id;"
rUlc.Open sql, DBCn, adOpenStatic, adLockReadOnly
If rUlc.BOF Then
 aktDC.ulcus = 2 ' kein Ulcus / 3 = nicht untersucht
Else
 aktDC.ulcus = rUlc!ulcus '0=oberflächlich, 1=tief
End If
' NaeUs AS Integer ' 0=1 Jahr, 1=6 Monate, 2=3 Monate
' Infekt AS Integer '0=nicht untersucht, 1=ja, 2=nein
' Deform AS Integer '0=nein, 1=ja
' Hyperk AS Integer '0=nein, 1=ja
' ZnUlcus AS Integer '0=nein, 1=ja
' ZnAmput AS Integer '0=nein, 1=ja

rfuss.Open "SELECT IF(nae_us like '%3%Mo%',2,if(nae_us like '%6%Mo%',1,0)) NaeUS, " & _
"if(wundinfektion like 'ja%',1,2) Infekt, " & _
"if(fußdeform in ('nein','keine','-',''),0,1) fußd, " & _
"(hyper_mEin RLIKE 'D[12345]' OR hyper_mEin like 'ja%') AND hyper_mEin<>fußdeform HypermEin, " & _
"not zn_ulcus in ('nein','','-') AND zn_ulcus<>fußdeform ZnUlcus, " & _
"zn_amput not in ('nein','','-') Amp, Fuß_ang " & _
"FROM fuss " & _
"WHERE pat_id = " & pid & _
" AND zn_amput in ('nein','ja') OR zn_amput<>fußdeform " & _
" AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM fuss WHERE pat_id = " & pid & " AND zeitpunkt BETWEEN qanf() AND qend() AND zeitpunkt > 20170717080000)", DBCn, adOpenStatic, adLockReadOnly
If Not rfuss.BOF Then
 aktDC.NaeUs = rfuss!NaeUs
 aktDC.Infekt = rfuss!Infekt
 aktDC.Deform = rfuss!Fußd
 aktDC.Hyperk = rfuss!hypermein
 aktDC.ZnUlcus = rfuss!ZnUlcus
 aktDC.ZnAmput = rfuss!amp
Else
 aktDC.NaeUs = 0
 aktDC.Infekt = 0
 aktDC.Deform = 0
 aktDC.Hyperk = 0
 aktDC.ZnUlcus = 0
 aktDC.ZnAmput = 0
End If

Set rUlc = Nothing
rUlc.Open "SELECT pat_id FROM ulcus WHERE pat_id = " & pid & _
" AND (beläge RLIKE 'eitr|schmier' OR exsudat RLIKE 'gelb|braun|stark|viel|zünd|mehr|eitr|eiter' OR " & _
" geruch RLIKE 'ja|röt|mittel|stark|äßig|vorhand' OR wundrand RLIKE 'röt|wärm|infiz|infekt|rot' OR " & _
" (wundumgebung RLIKE 'röt|wärm|infiz|infekt|rot|verfärb' AND not wundumgebung RLIKE 'keine Röt')) ", DBCn, adOpenStatic, adLockReadOnly
If Not rUlc.BOF Then aktDC.Infekt = 1

If aktDC.Infekt Or aktDC.ulcus < 2 Then aktDC.NaeUs = 2

Dim raltu As New ADODB.Recordset
raltu.Open "SELECT icd FROM diagnosen d WHERE pat_id = " & pid & " AND icd RLIKE '^L89.[234]' AND diagdatum < qanf() AND diagsicherheit<>'A' AND COALESCE(f6010,0)=0 ", DBCn, adOpenStatic, adLockReadOnly
If Not raltu.BOF Then aktDC.ZnUlcus = 1

Dim ranamp As New ADODB.Recordset
ranamp.Open "SELECT not ISNULL(amputation) AND amputation not in ('n','-','','entfällt','/','--','nn','u','.-') AND amputation NOT LIKE '?%' AND amputation not RLIKE '^n[ .]' AND amputation not RLIKE 'nein' amp FROM anamnesebogen a WHERE pat_id = " & pid, DBCn, adOpenStatic, adLockReadOnly
If Not ranamp.EOF And ranamp!amp = 1 Then aktDC.ZnAmput = 1
  
  
 If mitStr And ab317 Then TabPr "•  Pulsstatus: ", PulsStatus
 If mitStr And ab317 Then TabPr "Sensibilität: ", SensText
 If mitStr And ab317 Then
  Dim wrstr$
  wrstr = IIf(aktDC.Deform, "Fußdeformität, ", "") & _
  IIf(aktDC.Hyperk, "Hyperkeratose mit Einblutung, ", "") & _
  IIf(aktDC.ZnUlcus, "Z.n. Ulcus, ", "") & _
  IIf(aktDC.ZnAmput, "Z.n. Amputation, ", "") & _
  IIf(aktDC.Deform Or aktDC.Hyperk Or aktDC.ZnUlcus Or aktDC.ZnAmput, "ja", "nein")
  TabPr "Weiteres Risiko für Ulcus: ", wrstr
 End If
 If mitStr And ab317 Then TabPr "Ulcus: ", Switch(aktDC.ulcus = 0, "oberflächlich", aktDC.ulcus = 1, "tief", aktDC.ulcus = 2, "nein", 1, "nicht untersucht")
 If mitStr And ab317 Then TabPr "Wundinfektion: ", IIf(aktDC.Infekt = 0, "nicht untersucht", IIf(aktDC.Infekt = 1, "ja", "nein"))
 If mitStr And ab317 Then TabPr "Intervall der künftigen Fußinspektionen: ", IIf(aktDC.NaeUs = 0, "jährlich", IIf(aktDC.NaeUs = 1, "alle 6 Monate", "alle 3 Mo oder häufiger"))
 
 Dim merkmitstr%
 merkmitstr = mitStr
 If ab317 Then mitStr = 0
 Dim raFa As New ADODB.Recordset
 Dim diIsql$
 diIsql = "SELECT ICD, DiagSeite, DiagDatum FROM `diagnosen` WHERE pat_id = " & pid & " AND (icd like ""M14.6%"" Or icd like ""T79.%"" Or icd like ""L89.%"" Or icd like ""T89.%"" Or icd like ""T87.4%"" Or icd like ""Z44.1%"") AND diagsicherheit in (""G"",""V"",""Z"")  AND COALESCE(f6010,0)=0 " ' M14.6 = Charcot, M14.2 = Arthorpathie
 Dim lddat As Date
 raFa.Open "SELECT MAX(bhfb) AS lddat FROM `faelle` WHERE pat_id = " & pid, DBCn, adOpenStatic, adLockReadOnly
 If Not raFa.BOF Then
  If Not IsNull(raFa!lddat) Then
   lddat = raFa!lddat
   diIsql = diIsql & " AND (obdauer <> 0 Or diagdatum >= " & DatFor_k(lddat) & ")"
  End If
 End If
 Set raDT = Nothing
 Call raDT.Open(diIsql, DBCn, adOpenDynamic, adLockReadOnly)
 Dim obUlc%, FSauf$
 obUlc = obPosi(!Ulcera)
 If obUlc Or Not raDT.BOF Then
  aktDC.fußst = auff
  FSauf = "auffällig "
  If obUlc Then FSauf = FSauf & "(Anamnesebogen:Ulcera: " & !Ulcera & ") "
  Do While Not raDT.EOF
   Dim diags$
   If Not IsNull(raDT!DiagSeite) Then diags = raDT!DiagSeite
   FSauf = FSauf & "ICD: " & raDT!ICD & diags & " (" & Format$(raDT!DiagDatum, "d.m.yy") & ") "
   raDT.Move 1
  Loop
  If mitStr Then TabPr "Fußstatus:", FSauf
  Dim DruckNStSeite$
  If (trp And drp) And Not (tlp And dlp) Then
   aktDC.NStSeite = rE
   DruckNStSeite = "re (Pulse)"
  ElseIf Not (trp And drp) And (tlp And dlp) Then
   aktDC.NStSeite = li
   DruckNStSeite = "li (Pulse)"
  ElseIf VibrIKrep + VibrGZrep > VibrIKlip + VibrGZlip Then
   aktDC.NStSeite = rE
   DruckNStSeite = "re (Vibr)"
  ElseIf VibrIKrep + VibrGZrep < VibrIKlip + VibrGZlip Then
   aktDC.NStSeite = li
   DruckNStSeite = "li (Vibr)"
  ElseIf Mfrep > Mflip Then
   aktDC.NStSeite = rE
   DruckNStSeite = "re (Monofil.)"
  ElseIf Mfrep < Mflip Then
   aktDC.NStSeite = li
   DruckNStSeite = "li (Monofil.)"
  Else
   aktDC.NStSeite = gleich
   DruckNStSeite = "="
  End If
'  Enum DFSNiveau
'   St0%
'   St1%
'   St2%
'   AmpZeh%
'   St3%
'   AmpUS%
'   St4%
'   AmpOS%
'  End Enum

  If lies.obMySQL Then
   sql = "SELECT * FROM `diagnosen` WHERE (obdauer <> 0 Or (obdauer = 0 AND CONCAT(((month(diagdatum)+2) div 3)¡ year(diagdatum)) = '" & ZQuart(DokuDat) & "')) AND diagsicherheit<>'A' AND COALESCE(f6010,0)=0 AND pat_id = " & aktDC.Pat_id
  Else
   sql = "SELECT * FROM `diagnosen` WHERE (obdauer <> 0 Or (obdauer = 0 AND (int((month(diagdatum)+2) / 3) & year(diagdatum)) = '" & ZQuart(DokuDat) & "')) AND diagsicherheit<>'A' AND COALESCE(f6010,0)=0 AND pat_id = " & aktDC.Pat_id
  End If
  If lies.obMySQL Then
   sql = Replace$(sql, "¡", ",")
  Else
   sql = Replace$(Replace$(sql, "concat", ""), "¡", " & ")
  End If
  
'  Set rDT = aktDCb.OpenRecordset(sql) ' Set rDT = aktDCb.OpenRecordset("Diagnosen", dbOpenDynaset)
  Set raDT = Nothing
  raDT.Open sql, DBCn, adOpenDynamic, adLockReadOnly
  Dim aktDN As DFSNiveau
'  Dim maxAmp AS DFSNiveau
'  Dim aktDC.msei$ ' re, li
  Do While Not raDT.EOF
   aktDN = stNichts
   If raDT!ICD Like "Z44.1*" Then
    If InStrB(raDT!DiagText, "OS") <> 0 Then
     aktDN = AmpOS
    ElseIf InStrB(raDT!DiagText, "US") <> 0 Then
     aktDN = AmpUS
    ElseIf InStrB(raDT!DiagText, "Zeh") <> 0 Or InStrB(raDT!DiagText, "uß") <> 0 Then
     aktDN = AmpZeh
    End If
   End If
   If aktDN <> stNichts Then
    If aktDN > aktDC.maxAmp Then
     If (InStrB(raDT!DiagText, "re ") <> 0 Or Right$(raDT!DiagText, 2) = "re") Then aktDC.mSei = rE: aktDC.maxAmp = aktDN Else If (InStrB(raDT!DiagText, "li ") <> 0 Or Right$(raDT!DiagText, 2) = "li") Then aktDC.mSei = li: aktDC.maxAmp = aktDN
    End If
   End If
   raDT.Move 1
  Loop
'  sql = "SELECT * FROM `" & stftdb & "`.jpg WHERE pat_id = " & CStr(PID) & " AND cstr(int(month(patdatum)/3)+1) & year(patdatum) = """ & ZQuart(DokuDat) & """"
'  Set rafot = aktDCb.OpenRecordset(sql)
  
  Dim raFot As New ADODB.Recordset
#If False Then
'If lese.obMySQL Then
' If FtCn.State = 0 Then
'  Call acon(FotT)
' End If
' raFot.Open "SELECT * FROM Jpg WHERE pat_id = " & CStr(pid) & " AND CONCAT(floor(month(patdatum)+2) div 3,year(patdatum)) = """ & ZQuart(DokuDat) & """", FtCn, adOpenDynamic, adLockReadOnly
'Else
' raFot.Open "SELECT * FROM jpg WHERE pat_id = " & CStr(pid) & " AND cstr(int(month(patdatum)+2)/3) & year(patdatum) = """ & ZQuart(DokuDat) & """", CStrAcc & StFtDB, adOpenDynamic, adLockReadOnly
'End If
#Else
  raFot.Open "SELECT dokname FROM `dokumente` WHERE pat_id = " & pid & " AND dokart = 'jpg' AND dokname like '%Foto%' AND dokname like '% WA%'", DBCn, adOpenDynamic, adLockReadOnly
#End If

  Dim maxDN As DFSNiveau, aktWA$, aktSei As DMPSeite
  Do While Not raFot.EOF
   Dim wa$, körperteil$, beschreibung$, p1&
   wa = Left(Trim$(Mid$(raFot!DokName, InStr(raFot!DokName, " WA") + 3)), 2)
   körperteil = Mid$(raFot!DokName, InStr(raFot!DokName, "Foto ") + 5)
   For p1 = 1 To Len(körperteil) - 1
    If IsNumeric(Mid$(körperteil, p1, 1)) And (Mid$(körperteil, p1 + 2, 1) = "." Or Mid$(körperteil, p1 + 1, 1) = ".") Then
     beschreibung = Mid$(körperteil, InStr(p1, körperteil, " ") + 1)
'     beschreibung = trim$(left(beschreibung, InStr(beschreibung, "WA") - 1))
     Dim poswa%
     poswa = InStr(beschreibung, "WA ")
     If poswa = 0 Then poswa = InStr(beschreibung, "Wagner ")
     If poswa > 0 Then
      beschreibung = Left(beschreibung, poswa - 1)
     End If
     beschreibung = Trim$(beschreibung)
     körperteil = Trim$(Left(körperteil, p1 - 1))
     Exit For
    End If
   Next p1
   aktDN = stNichts
    If Not IsNull(wa) Then
     If Len(wa) > 0 Then
      Select Case Left(wa, 1)
       Case "1": aktDN = St0: aktWA = wa
       Case "2": aktDN = St1: aktWA = wa
       Case "3": aktDN = St2: aktWA = wa
       Case "4": aktDN = St3: aktWA = wa
'       Case Else: aktdn = Stnichts ' Schätzung
      End Select
     End If
    End If
   If aktDN <> stNichts Then
    If aktDN > maxDN Then
     If (InStrB(körperteil, "re ") <> 0 Or Right$(körperteil, 2) = "re") Then
        If aktDN > aktDC.maxAmp Then aktDC.mSei = rE
        maxDN = aktDN
        If aktWA > aktDC.mWA Then aktDC.mWA = aktWA
     ElseIf (InStrB(körperteil, "li ") <> 0 Or Right$(körperteil, 2) = "li") Then
        If aktDN > aktDC.maxAmp Then aktDC.mSei = li
        maxDN = aktDN
        If aktWA > aktDC.mWA Then aktDC.mWA = aktWA
     ElseIf (InStrB(beschreibung, "re ") <> 0 Or Right$(beschreibung, 2) = "re") Then
        If aktDN > aktDC.maxAmp Then aktDC.mSei = rE
        maxDN = aktDN
        If aktWA > aktDC.mWA Then aktDC.mWA = aktWA
     ElseIf (InStrB(beschreibung, "li ") <> 0 Or Right$(beschreibung, 2) = "li") Then
        If aktDN > aktDC.maxAmp Then aktDC.mSei = li
        maxDN = aktDN
        If aktWA > aktDC.mWA Then aktDC.mWA = aktWA
     End If
    End If
   End If
   raFot.Move 1
  Loop ' While Not raFot.EOF
  
'  If aktDC.mwa = "" Then
  If Not raDT.BOF Then
   raDT.MoveFirst
   aktSei = unbek
   Do While Not raDT.EOF
    If InStrB(raDT!ICD, "L89.") <> 0 Then
     aktWA = MAX(Mid$(raDT!ICD, 5, 1) - 1, 0) ' Korrektur 3.1.15 z.B. L89.09
     If aktWA > aktDC.mWA Then
      aktDC.mWA = aktWA
      If IsNull(raDT!DiagSeite) Then
       aktSei = unbek
      Else
       Select Case UCase$(raDT!DiagSeite)
        Case "R": aktSei = rE
        Case "L": aktSei = li
        Case "B": aktSei = gleich
        Case Else: aktSei = unbek
       End Select
      End If
     End If
    End If
    raDT.Move 1
   Loop
  End If ' Not raDT.BOF Then
  Set raDT = Nothing
  If Len(aktDC.mWA) = 1 Then ' 5.10.08
   Dim ArmSt$
   If aktDC.mWA = "0" Then
    If aktDC.Puls = unauff Then ArmSt = "A" Else ArmSt = "C" ' 0 eher nicht entzündet
   ElseIf aktDC.mWA > "0" Then
    If aktDC.Puls = unauff Then ArmSt = "B" Else ArmSt = "D"
   Else
   ' WagSt = ""
   End If
   aktDC.mWA = aktDC.mWA & ArmSt
  End If
'  End If
  Dim DruckmSei$
  If aktDC.mSei = unbek Then aktDC.mSei = aktSei
  Select Case aktDC.mSei
   Case rE: DruckmSei = "rechts"
   Case li: DruckmSei = "links"
   Case gleich: DruckmSei = "gleich"
   Case unbek: DruckmSei = vNS
  End Select
  If mitStr Then TabPr "Schwerer betr. Fuß:", IIf(LenB(DruckmSei) = 0, DruckNStSeite, DruckmSei)
  If aktDC.mSei = gleich Or aktDC.mSei = unbek Then
   aktDC.mSei = aktDC.NStSeite
   If aktDC.mSei = gleich Or aktDC.mSei = unbek Then aktDC.mSei = rE ' meistens rechts
  End If
  If aktDC.mWA > "" Then
   If mitStr Then TabPr "Grad n. Wagner+Armstrong:", aktDC.mWA
  End If
  Set raFot = Nothing
'#If False Then
'  Dim posi%
'  posi = InStr(!Diagnosen, "Wagner")
'  Dim inf AS Boolean, isch AS Boolean, ArmSt$
'  inf = False
'  isch = False
'  If diI("T79.") OR diI("T89") OR diI("T87.4") Then inf = True
'  'If diI("M12") OR diI("M20") OR diI("M21") OR diI("M77.3") OR diI("M84") OR diI("M84.17") Then def = True
'  If (trp AND drp) Or (tlp AND dlp) Or (trp AND dlp) Or (tlp AND drp) Then isch = True
'  ArmSt = "A"
'  If inf Then
'   ArmSt = "B"
'   If isch Then ArmSt = "D"
'  ElseIf isch Then
'   ArmSt = "C"
'  End If
'  If posi <> 0 Then
'   if mitstr then TabPr "Grad nach Wagner und Armstrong: ", IIf(posi <> 0, Mid$(!Diagnosen, posi + 7, 1) + ArmSt, "0A")
'  End If
'#End If
#If vorca2008 Then
  Set raDT = Nothing
  Call raDT.Open("SELECT * FROM `diagnosen` WHERE pat_id = " & pid & " AND icd like '" & "M14.6%" & "' AND diagsicherheit in (""G"",""V"") AND COALESCE(f6010,0)=0 ", DBCn, adOpenDynamic, adLockReadOnly)
  If Not raDT.BOF Then
   Dim Seite$
   If Not IsNull(raDT!DiagSeite) Then
    Seite = raDT!DiagSeite
    Select Case raDT!DiagSeite
     Case "R": aktDC.oap = rE
     Case "L": aktDC.oap = li
     Case "B": aktDC.oap = gleich ' heißt hier: unbekannt
     Case Else: aktDC.oap = unbek ' heißt hier: keine
    End Select
   Else
    Seite = "ja"
    aktDC.oap = gleich
   End If
   If mitStr Then TabPr "Osteoarthropathie ", Seite
  Else
   aktDC.oap = unbek ' heißt hier: keine
   If mitStr Then TabPr "Osteoarthropathie:", "keine"
  End If
  Set raDT = Nothing
#End If ' vorca2008

 ElseIf tfeld(!Ulcera) = "" And tfeld(![Puls Atp]) = "" And tfeld(![Puls Adp]) = "" And tfeld(![Vibration IK]) = "" And tfeld(![Vibration Großzehe]) = "" And tfeld(![Kalt-Warm]) = "" Then
  aktDC.fußst = ndok
  If mitStr Then TabPr "Fußstatus:", "" '"nicht dokumentiert"
 Else ' obUlc Or Not raDT.BOF Then
  aktDC.fußst = unauff
  If mitStr Then TabPr "Fußstatus:", "unauffällig"
 End If ' obUlc Or Not raDT.BOF Then
 mitStr = merkmitstr
 
 If aktDC.dtyp = "1" Then
  If aktDC.PrRR <> "" Then
   If mitStr Then TabPr "Blutdruck:", aktDC.PrRR
  End If
  If mitStr Then TabPr "HbA1c:", DMPHbA1c
  If mitStr Then TabPr "Serum-Kreatinin:", DMPCrea
  If mitStr Then TabPr "Urin-Albumin:", DMPUAlb
 End If
 
#If vorca2008 Then
' Dim oblaser AS Boolean
 Dim kontrah$
 aktDC.oblaser = True
 If Not obPosi(![Netzhaut gelasert]) Then
  aktDC.oblaser = False
'  rEi.Seek "=", PID
  Set raEi = Nothing
  Call raEi.Open("SELECT * FROM `eintraege` WHERE pat_id = " & pid, DBCn, adOpenDynamic, adLockReadOnly)
  If Not raEi.BOF Then
   Do
    Dim StartPosi&
    If raEi.EOF Then Exit Do
    kontrah = Replace$(raEi!Inhalt, " ", "")
    StartPosi = 1
    StartPosi = InStr(kontrah, "Netzhautschongelasert")
    If StartPosi > 1 Then StartPosi = StartPosi + 20 ' ca.
    If StartPosi = 0 Then StartPosi = 1
    If InStr(StartPosi, kontrah, "laser") <> 0 Then ' 16.8.15: Instr statt InstrB wg. Startposi (müsste sonst ca. doppelt genommen werden
     aktDC.oblaser = True
     Exit Do
    End If
    raEi.Move 1
   Loop
  End If
 End If
' Handlungsbedarf: spätere Lasertherapie
If aktDC.dtyp = "2" Then If mitStr Then TabPr "Lasertherapie:", IIf(aktDC.oblaser, "ja", "nein")
#End If 'vorca2008
 
  If mitStr And ab315 And Not ab317 Then
   TabPr "•  Injektionsstellen:", IIf(aktDC.inj = 0, "nicht untersucht", IIf(aktDC.inj = 1, "unauffällig", "auffällig"))
  End If
  
' 3. Reiter
 If diI("Z49", , QAnf(ZQuart(DokuDat))) Then
  aktDC.neuDial = True
  Set raDT = Nothing
 End If
 If diI("H54", , QAnf(ZQuart(DokuDat)), , "diagseite in (' ','B')") Then
  aktDC.neuErbl = True
  Set raDT = Nothing
 End If
 If diI("Z44", , QAnf(ZQuart(DokuDat)), True) Then
  aktDC.neuAmp = True
  Set raDT = Nothing
 End If
 If diI("I21", , QAnf(ZQuart(DokuDat)), True) Then
  aktDC.neuMI = True
  Set raDT = Nothing
 End If
 If diI("I61 I62 I63 I64", , QAnf(ZQuart(DokuDat)), True) Then
  aktDC.neuApo = True
  Set raDT = Nothing
 End If
 
 If aktDC.dtyp = "1" Then If mitStr Then TabPr "Lasertherapie:", IIf(aktDC.oblaser, "ja", "nein")
  If ab315 Then
   If mitStr Then
    If aktDC.FEn(8) Then ereig = "Nierenersatzth." & IIf(aktDC.FE(8), "(bek)", "(ICD " & aktDC.FEi(8) & " vor " & Format(aktDC.FEd(8), "d.m.yy") & ")") & ", "
    If aktDC.FEn(10) Then ereig = ereig & "Erblindung" & IIf(aktDC.FE(10), "(bek)", "(ICD " & aktDC.FEi(10) & " vor " & Format(aktDC.FEd(10), "d.m.yy") & ")") & ", "
    If aktDC.FEn(13) Then ereig = ereig & "Amputation" & IIf(aktDC.FE(13), "(bek)", "(ICD " & aktDC.FEi(13) & " vor " & Format(aktDC.FEd(13), "d.m.yy") & ")") & ", "
    If aktDC.FEn(4) Then ereig = ereig & "Herzinfarkt" & IIf(aktDC.FE(4), "(bek)", "(ICD " & aktDC.FEi(4) & " vor " & Format(aktDC.FEd(4), "d.m.yy") & ")") & ", "
    If aktDC.FEn(5) Then ereig = ereig & "Schlaganfall" & IIf(aktDC.FE(5), "(bek)", "(ICD " & aktDC.FEi(5) & " vor " & Format(aktDC.FEd(5), "d.m.yy") & ")") & ","
    If LenB(ereig) = 0 Then ereig = "keine d.gen." Else ereig = Left$(ereig, Len(ereig) - 2)
    TabPr "relev.Ereignisse: ", ereig
   End If
  End If
 
 If raAna!Vorgestellt > DokuDat - 92 Then ' Anamnese nur ein Quartal berücksichtigen
  Dim hypoStr$
  aktDC.hypoZAn = 0
  If obPosi(![Fremde Hilfe pa]) Then
   If IsNumeric(![Fremde Hilfe pa]) Then
    aktDC.hypoZAn = ![Fremde Hilfe pa]
    aktDC.hypoZ = aktDC.hypoZAn
   End If
   If mitStr Then hypoStr = CStr(![Fremde Hilfe pa])
  Else
   If mitStr Then hypoStr = IIf(LenB(tfeld(![Fremde Hilfe pa])) = 0, "keine Angaben", "keine")
  End If
  If mitStr Then TabPr "Schw.Hypoglyk./12 Mon:", hypoStr
 Else
  rs.Open "SELECT COUNT(0) AS ct FROM `eintraege` WHERE pat_id = " & pid & " AND art = 'hypo' AND zeitpunkt >= " & DatFor_k(QAnf(ZQuart(DokuDat))), DBCn, adOpenStatic, adLockReadOnly
  aktDC.hypoZKK = rs!ct
  Set rs = Nothing
  If IsNumeric(aktDC.dmpHypos) Then If aktDC.dmpHypos <> 0 Then aktDC.hypoZ = aktDC.dmpHypos
  aktDC.hypoZ = MAX(aktDC.hypoZ, aktDC.hypoZKK)
  If aktDC.hypoZ Then
   hypoStr = aktDC.hypoZ
  Else
   hypoStr = "keine"
   rs.Open "SELECT GROUP_CONCAT(replace(inhalt,'kommt ','komme ')) FROM `eintraege` WHERE pat_id = " & pid & " AND art = 'uzu' AND zeitpunkt >= " & DatFor_k(QAnf(ZQuart(DokuDat))), DBCn, adOpenStatic, adLockReadOnly
   If Not rs.BOF Then
    hypoStr = hypoStr & " ('Unterzucker?': " & rs.Fields(0) & ")"
   End If
   Set rs = Nothing
  End If
  If mitStr Then TabPr "Schw.Hypoglyk./Q.:", hypoStr
 End If
 
 If raAna!Vorgestellt > DokuDat - 92 Then ' Anamnese nur ein Quartal berücksichtigen
  Dim khsStr$
' Handlungsbedarf: Krankenhauseinweisungen
  aktDC.krZAn = 0
  If obPosi(![keto pa]) Then
   If IsNumeric(![keto pa]) Then
'   aktDC.krZAn = ![keto pa]
    If IsDate(raAna![keto pa]) Then
     aktDC.krZAn = 1
    ElseIf IsNumeric(raAna![keto pa]) Then
     aktDC.krZAn = raAna![keto pa]
    End If
    aktDC.krz = aktDC.krZAn
   End If
   If mitStr Then khsStr = CStr(![keto pa])
  Else
   If mitStr Then khsStr = IIf(LenB(tfeld(![Fremde Hilfe pa])) = 0, "keine Angaben", "keine")
  End If
  If mitStr Then TabPr "Krankenhausaufenthalte / Ketoazidosen /12 Mon:", khsStr
 Else
  raKH.Open "SELECT * FROM kheinweis WHERE pat_id = " & pid & " AND zeitpunkt > " & DatFor_k(DokuDat - 92), DBCn, adOpenDynamic, adLockReadOnly
  If Not raKH.EOF Then
   aktDC.khew = True
   If mitStr Then TabPr "Krankenhauseinweisung", "(am " + Format$(raKH!Zeitpunkt, "dd/mm/yy") + IIf(tfeld(raKH!Ziel) <> "nkenhauseinweisung", " nach:" + raKH!Ziel, "") + _
             " mit: " + tfeld(raKH!Diagnose) + ")"
  End If
  Set raKH = Nothing
' nur ein Datensatz 2008
  rs.Open "SELECT COUNT(0) AS ct FROM `eintraege` WHERE pat_id = " & pid & " AND art = 'kra' AND zeitpunkt >= " & DatFor_k(QAnf(ZQuart(DokuDat))), DBCn, adOpenStatic, adLockReadOnly
  aktDC.krZKK = rs!ct
  Set rs = Nothing
  If (IsNumeric(aktDC.dmpKhsA)) Then If aktDC.dmpKhsA <> 0 Then aktDC.krZKK = aktDC.dmpKhsA
  aktDC.krz = aktDC.krZKK
 End If
 
 
' 4. Reiter
 Dim TherArtAkt$
' TherArtAkt = TherArtAkt
' TherArtAkt = therart_erm(aktDC.Pat_id, 0)
 TherArtAkt = "Diät"
 On Error Resume Next
 TherArtAkt = DBCn.Execute("SELECT therart FROM `therarten` tha WHERE pat_id=" & pid & " AND DATE(zp)<= " & Format(Now(), "yyyymmdd") & " ORDER BY zp DESC,mpnr DESC LIMIT 1").Fields(0)
 On Error GoTo fehler
 
 Dim thLang$
 aktDC.tart = ThaRang(TherArtAkt, thLang) ' Therapieart
 
' SELECT CASE TherArt(aktDC.Pat_id, 0)
'  Case "Diät?":  aktDC.tart = diät
'  Case "": aktDC.tart = unbek
'  Case "(I)CT?", "I/CT?", "ICT": aktDC.tart = ict
'  Case "CSII":   aktDC.tart = CSII
'  Case "CT":     aktDC.tart = ct
'  Case "Komb":   aktDC.tart = komb
'  Case "OAD":     aktDC.tart = oad
' End SELECT
' If aktDC.tart = csii Then
 If ab315 Then
  If mitStr Then TabPr "•  Insulin oder -analoga:", IIf(aktDC.obIns = adja Or aktDC.obAnal = adja, "ja", "nein")
 Else
  If mitStr Then TabPr "Insulin:", IIf(aktDC.obIns, IIf(aktDC.insz > 2, IIf(Not ab315 And aktDC.VorM(5) = 1, "(unverändert) ", "  ") & "ja, intensiviert", IIf(Not ab315 And aktDC.VorM(5) = 2, "(unverändert) ", "  ") & "ja, " & TherArtAkt), IIf(Not ab315 And aktDC.VorM(5) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "Nein")
  If mitStr Then TabPr "Insulin-Analoga:", IIf(aktDC.obAnal, IIf(aktDC.insz > 2, IIf(Not ab315 And aktDC.VorM(6) = 1, "(unverändert) ", "  ") & "ja, intensiviert", IIf(Not ab315 And aktDC.VorM(6) = 2, "(unverändert) ", "  ") & "ja, " & TherArtAkt), IIf(Not ab315 And aktDC.VorM(6) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "Nein")
 End If
 
 Dim metText$
 aktDC.obNI = (aktDC.eGFR < 45)
 
' If mitStr Then DmPStrS = DmPStrS & vbCrLf
 If aktDC.dtyp <> "1" Then
  If mitStr Then TabPr "Glibenclamid:", IIf(ab315, IIf(aktDC.obGlib, "ja", IIf(aktDC.obNI, "Kontraindikation", "nein")), IIf(aktDC.obGlib, IIf(Not ab315 And aktDC.VorM(0) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "ja", IIf(Not ab315 And aktDC.VorM(0) > 0, "(unverändert) ", "  ") & IIf(aktDC.obNI, "Kontraindikation", "Nein")))
  If Not aktDC.obGlib And aktDC.obNI Then aktDC.obGlib = adki
  If aktDC.obmetf Then
   metText = "  Ja"
   If Not ab315 And aktDC.VorM(1) = 0 And aktDC.obVorb Then metText = "(unverändert)" & metText
   If aktDC.obNI Then metText = metText & ": Sollte evtl. aufgrund Creatininerhöhung abgesetzt werden"
  Else
   metText = IIf(aktDC.obNI, "Kontraindikation", "Nein")
   If Not ab315 Then If aktDC.VorM(1) = 1 Or aktDC.VorM(1) = 2 Then metText = "(unverändert)" & metText
  End If
  If Not aktDC.obmetf And aktDC.obNI Then aktDC.obmetf = adki
  If mitStr Then TabPr "Metformin:", metText
  If ab315 Then
   If mitStr Then TabPr "Sonst.or.antidiab.Medik.:", IIf(aktDC.obGlucI = adja Or aktDC.obSHGlin = adja Or aktDC.obGlit = adja Or aktDC.obsglt2 = adja Or aktDC.obdpp4 = adja Or aktDC.obglp1 = adja Or aktDC.obSonstAD = adja, "Ja", "Nein")
  Else
   If mitStr Then TabPr "Glucosidase-Inhib.:", IIf(aktDC.obGlucI, IIf(Not ab315 And aktDC.VorM(2) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "Ja", IIf(Not ab315 And aktDC.VorM(2) > 0, "(unverändert) ", "  ") & "Nein")
   If mitStr Then TabPr "Sonst. Sulf'hst./Glinide:", IIf(aktDC.obSHGlin, IIf(Not ab315 And aktDC.VorM(3) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "Ja", Not ab315 And IIf(aktDC.VorM(3) > 0, "(unverändert) ", "  ") & "Nein")
   If aktDC.obGlit = adnein And diI("I50") Then aktDC.obGlit = adki
   If mitStr Then TabPr "Glitazone:", IIf(aktDC.obGlit = adja, IIf(Not ab315 And aktDC.VorM(4) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "Ja", IIf(Not ab315 And aktDC.VorM(4) > 0, "(unverändert) ", "  ") & IIf(aktDC.obGlit = adki, "Kontraindikation", "Nein"))
   If mitStr Then TabPr "Sonstige Antidiab.:", IIf(aktDC.obSonstAD = adja Or aktDC.obdpp4 = adja Or aktDC.obglp1 = adja Or aktDC.obsglt2 = adja, IIf(Not ab315 And aktDC.VorM(12) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "Ja", IIf(Not ab315 And aktDC.VorM(12) > 0, "(unverändert) ", "  ") & "Nein")
  End If
 End If ' aktDC.dtyp <> "1"
 If mitStr Then TabPr "Thrombozytenhemmer:", IIf(aktDC.obThro, IIf(Not ab315 And aktDC.VorM(9) = 1, "(unverändert) ", "  ") & "ja", IIf(Not ab315 And aktDC.VorM(9) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "nein") & IIf(aktDC.obOAK, "; orale Antikoagul.", "")
 If mitStr Then TabPr "ACE-Hemmer:", IIf(aktDC.obACEH, IIf(Not ab315 And aktDC.VorM(11) = 1, "(unverändert) ", "  ") & "ja", IIf(Not ab315 And aktDC.VorM(11) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & IIf(aktDC.obAT1, "AT1", "nein"))
 If mitStr Then TabPr "Betablocker:", IIf(aktDC.obBetabl, IIf(Not ab315 And aktDC.VorM(10) = 1, "(unverändert) ", "  ") & "ja", IIf(Not ab315 And aktDC.VorM(10) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "nein")
' If mitStr Then DmPStrS = DmPStrS & vbCrLf
 If mitStr Then TabPr "HMG-CoA-Red'-Hemmer:", IIf(aktDC.obHMG, IIf(Not ab315 And aktDC.VorM(7) = 1, "(unverändert) ", "  ") & "ja", IIf(Not ab315 And aktDC.VorM(7) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "nein")
 If ab315 Then
  If mitStr Then TabPr "Diuretika:", IIf(aktDC.obDiur, IIf(Not ab315 And aktDC.VorM(13) = 1, "(unverändert) ", "  ") & "ja", IIf(Not ab315 And aktDC.VorM(13) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "nein")
 Else
  If mitStr Then TabPr "Antihypertensiva:", IIf(aktDC.obAntihyp, IIf(Not ab315 And aktDC.VorM(8) = 1, "(unverändert) ", "  ") & "ja", IIf(Not ab315 And aktDC.VorM(8) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "nein")
 End If

' 5. Reiter
' If mitStr Then DmPStrS = DmPStrS & vbCrLf
 If Not ab315 Then If mitStr Then TabPr "•  Schulungen vor DMP:", IIf(aktDC.SW(0) > 0, " D.m.: " & Format$(aktDC.SW(0), "d.m.yy"), "") & IIf(aktDC.SW(1) > 0, " Hypert: " & Format$(aktDC.SW(1), "d.m.yy"), "") & IIf(aktDC.SW(2) > 0, " keine: " & Format$(aktDC.SW(2), "d.m.yy"), "")
 
' Schulungen empfohlen
' Set raLau = Nothing
' raLau.Open "SELECT zeitpunkt,CAST(wert AS decimal(10,1)) wert FROM (" & lsql & ") AS innen WHERE abkü RLIKE '^HBA[1C]' AND CAST(wert AS decimal) < 22 ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly ' wegen falsch eingetragener Fremdlabore gestrichen: AND einheit = '%'
' Set raLau = hollabor(Pat_id, "^HBA[1C]", 0, 22)
 If lwZahl Then
  alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
  Labs = LabPat(LA_HbA1c, Pat_id)
  Dim zpd1$, zpdL$, zpr1$, zprL$
' If Not raLau.BOF Then
  If Labs.Abkü <> "" Then
   If Labs.WertSg <> vNS Then
    pos = InStr(Labs.WertSg, ">") ' Pat_id 744
    If pos > 0 Then Labs.WertSg = Mid(Labs.WertSg, pos + 1)
    If Labs.WertSg >= 8# Then
     DMSchulz = SchulzBest(!Pat_id, zpd1, zpdL, DokuDat - 2 * 365)
     If ((aktDC.tart = OAD Or aktDC.tart = offen Or aktDC.tart = Diät) And DMSchulz < 4) Or DMSchulz < 6 Then
      If zplschul(Pat_id) < DokuDat - 92 Then
'    Debug.Print Pat_id
       If raAna!GebDat > DokuDat - 365 * 83 Then ' Schulung nicht mehr unbedingt über 83 Jahre empfehlen
        aktDC.SE(0) = DokuDat ' SE(2) Schulung empfohlen
       End If
      End If
     End If
    End If
   End If
  End If
 End If ' lwzahl
 If aktDC.dmpDMSchulEmpf = "j" Then
  On Error Resume Next
  aktDC.SE(0) = CDate(DBCn.Execute("SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = " & CStr(pid) & " AND art = 'dmperg' AND instr(inhalt,'D.m.-Schul.empf.: j')").Fields(0))
  On Error GoTo fehler
  If aktDC.SE(0) = 0 Then aktDC.SE(0) = DokuDat - 1
 ElseIf aktDC.dmpDMSchulEmpf = "n" Then
  aktDC.SE(0) = 0
 End If
 
 If aktDC.dmpHypertSchulEmpf = "j" Then
  aktDC.SE(1) = 0
  On Error Resume Next
  aktDC.SE(1) = CDate(DBCn.Execute("SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = " & CStr(pid) & " AND art = 'dmperg' AND instr(inhalt,'Hypert-Schul.empf.: j')").Fields(0))
  On Error GoTo fehler
  If aktDC.SE(1) = 0 Then aktDC.SE(1) = DokuDat - 1
 ElseIf aktDC.dmpHypertSchulEmpf = "n" Then
  aktDC.SE(1) = 0
 End If
 
 If Not ab315 Then
  If mitStr Then TabPr "D.m.-Schul.empfohl:", IIf(aktDC.SE(0) > 0, " D.m. " & Format$(aktDC.SE(0), "d.m.yy"), "") & IIf(aktDC.SE(1) > 0, " Hypert: " & Format$(aktDC.SE(1), "d.m.yy"), "") & IIf(aktDC.SE(2) > 0, " keine: " & Format$(aktDC.SE(2), "d.m.yy"), "")
  If mitStr Then TabPr "DMP-Schulung D.m.:", IIf(aktDC.DS(0) > 0, " ja: " & Format$(aktDC.DS(0), "d.m.yy"), "") & IIf(aktDC.DS(1) > 0, " nein: " & Format$(aktDC.DS(1), "d.m.yy"), "") & IIf(aktDC.DS(2) > 0, " nicht mögl: " & Format$(aktDC.DS(2), "d.m.yy"), "")
  If mitStr Then TabPr "DMP-Schulung Hypt:", IIf(aktDC.HS(0) > 0, " ja: " & Format$(aktDC.HS(0), "d.m.yy"), "") & IIf(aktDC.HS(1) > 0, " nein: " & Format$(aktDC.HS(1), "d.m.yy"), "") & IIf(aktDC.HS(2) > 0, " nicht mögl: " & Format$(aktDC.HS(2), "d.m.yy"), "")
  If aktDC.kSE <> 0 Then If mitStr Then TabPr " Zuletzt k.Sch.empf:", Format$(aktDC.kSE, "d.m.yy")
  If mitStr Then
   If obPosi(![Jahr letzte Diabetesschulung]) Then
    Dim OrtS$
    If Not IsNull(![Ort Schulung]) Then OrtS = " (" + ![Ort Schulung] + ")"
    SchulStr = ![Jahr letzte Diabetesschulung] + OrtS
   Else
    SchulStr = "keine"
   End If
   TabPr "Schulung vor Zuweisung:", SchulStr
  End If
 End If
 
 If mitStr Then
  Dim empfstr$
  empfstr = IIf(aktDC.dmpDMSchulEmpf = "j", "Diabetes, ", "") & IIf(aktDC.dmpHypertSchulEmpf = "j", "Hypertonie, ", "") & IIf(aktDC.dmpDMSchulEmpf <> "j" And aktDC.dmpHypertSchulEmpf <> "j", "keine, ", "")
  empfstr = Left$(empfstr, Len(empfstr) - 2)
  TabPr "•  Schulungen empfohlen: ", empfstr

  empfstr = ""
  Select Case aktDC.dmpDMSchulWahrg
   Case "j"
    empfstr = "ja"
   Case "n"
    empfstr = "nein"
   Case "n mögl"
    empfstr = "nicht möglich"
   Case Else ' "u", "n empf"
    empfstr = "b.d.letzten Doku n.empf."
  End Select
  TabPr "Diabetesschulung   wahrg.: ", empfstr
  
  empfstr = ""
  Select Case aktDC.dmpHypertSchulWahrg
   Case "j"
    empfstr = "ja"
   Case "n"
    empfstr = "nein"
   Case "n mögl"
    empfstr = "nicht möglich"
   Case Else ' "u", "n empf"
    empfstr = "b.d.letzten Doku n.empf."
  End Select
  TabPr "Hypertonieschulung wahrg.: ", empfstr
 End If
 
 Dim sqls$
 sqls$ = _
 "SELECT `leistungen`.pat_id AS pat_id, `leistungen`.zeitpunkt AS zeitpunkt, ucase(`leistungen`.leistung) AS lst" & _
 " FROM `leistungen` INNER JOIN `faelle` ON `faelle`.pat_id = `leistungen`.pat_id AND `faelle`.quartal = cstr(int(month(`leistungen`.zeitpunkt)+2) / 3) & year(`leistungen`.zeitpunkt) " & _
 " WHERE `faelle`.pat_id = " & CStr(pid)
 If lies.obMySQL Then sqls = Replace$(Replace$(Replace$(sqls, "int(", "floor("), "cstr", ""), " / ", " div ")
' Set lpp = aktDCb.OpenRecordset(sqls, dbOpenDynaset)
 Call lapp.Open(sqls, DBCn, adOpenDynamic, adLockReadOnly)
 DMSchulz = 0
 HSchulz = 0
 zpd1 = Format$(DokuDat + 1, "d.m.yy")
 zpdL = "1.1.1900"
 zpr1 = Format$(DokuDat + 1, "d.m.yy")
 zprL = "1.1.1900"
 If Not lapp.BOF Then
  lapp.MoveFirst
  Do While Not lapp.EOF
   Call SchulungszifferZuord(lapp!lst, DMSchulz, HSchulz, lapp!Zeitpunkt, zpd1, zpdL, zpr1, zprL)
   lapp.Move 1
  Loop
 End If
 
 DMSchulz = SchulzBest(!Pat_id, zpd1, zpdL)
 If mitStr Then
  SchulStr = IIf(DMSchulz > 1, CStr(DMSchulz) + " (Diabetes: " + zpd1 + "-" + zpdL + ")", "") + IIf(HSchulz > 1, CStr(HSchulz) + " (Hypertonie: " + zpr1 + "-" + zprL + ")", "")
  If SchulStr = "" Then SchulStr = "keine"
  TabPr "Bish.Schulg.b.uns (ges.):", SchulStr
'  DmPStrS = DmPStrS & vbCrLf
 End If
 
 
' 6. Reiter
 If ab315 Then
  If mitStr Then
   Dim infostr$
   If aktDC.dmpKKTabakEmpf = "j" Or aktDC.dmpKKkTrainEmpf = "j" Or aktDC.dmpKKErnEmpf = "j" Then
    If aktDC.dmpKKTabakEmpf = "j" Then infostr = "Tabakverzicht, "
    If aktDC.dmpKKErnEmpf = "j" Then infostr = "Ernährungsberatung, "
    If aktDC.dmpKKkTrainEmpf = "j" Then infostr = "körperl.Training, "
    infostr = Left$(infostr, Len(infostr) - 2)
   Else
    infostr = "nichts angegeben"
   End If
   TabPr "•  Info'angeb.KK gewü.: ", infostr
  End If
 Else
  If aktDC.dtyp = "2" Then
   If mitStr Then TabPr "  Aufgabe Tabak empf.:", IIf(aktDC.Tabak, "ja", "nein")
   If !Größe <> 0 Then
''   if mitstr then TabPr "Ernährungsber. empf.:", IIf(CDbl(replace$(!`bmi`, ".", ",")) > 24.9, "ja", "nein")
'    Dim bmi!
'    bmi = !Gewicht * IIf(!Gewicht < 3, 100, 1) / !Größe / !Größe '* 10000
'    Do
'     If bmi = 0 Then Exit Do
'     If bmi > 8 Then Exit Do
'     bmi = bmi * 10
'    Loop
'    If bmi > 24.9 Then aktDC.ernb = True
    If aktDC.bmi >= 25 Then aktDC.ernb = True
    If mitStr Then TabPr "Ernährungsber. empf.:", IIf(aktDC.ernb, "ja", "nein")
   End If
  End If
 End If
 
 Dim grenze#
 If raAna!Palter >= 75 Then
  grenze = 7.5
 ElseIf raAna!Palter >= 70 Then
  grenze = 7
 ElseIf raAna!Palter >= 65 Then
  grenze = 6.8
 Else
  grenze = 6.5
 End If
 aktDC.hbEmpf = IIf(aktDC.bekHb > grenze, senken, halten)
 If IsNumeric(aktDC.dmpHbA1cZiel) Then
   If aktDC.bekHb <= CDbl(aktDC.dmpHbA1cZiel) Then
    aktDC.hbEmpf = halten
   Else
    aktDC.hbEmpf = halten
   End If
 End If
 If ab315 Then
  If mitStr Then TabPr "HbA1c-Zielwert: ", IIf(aktDC.hbEmpf = halten, "erreicht", "noch nicht erreicht")
 Else
  If mitStr Then TabPr "Zielvereinb. HbA1c:", IIf(aktDC.hbEmpf = senken, "senken", "halten")
  Dim rrEmpfehlung$
  rrEmpfehlung = rrEmpf(aktDC.RRsyst, aktDC.RRdiast, CStr(pid))
  aktDC.rrEmpf = IIf(rrEmpfehlung = "halten", halten, senken)
  If mitStr Then TabPr "Zielvereinb. Blutdruck:", rrEmpfehlung ' manuell auszufüllen
 End If
 
 If ab315 Then
  If mitStr Then TabPr "Bhdlg.Fußeinr.: ", IIf(aktDC.fußst = auff Or aktDC.dmpUewFuss = "j", "ja", "nein")
  If mitStr Then TabPr "Diab'bez.stat.Einweisung: ", IIf(aktDC.dmpEinwDM = "j", "ja", "nein")
 End If
 
 Dim AugU$, AugUDat As Date
 Dim rBr As New ADODB.Recordset
 AugU = vNS
' Set rdo = aktDCb.OpenRecordset("SELECT * FROM `" + QMdbAkt + "`.`dokumente` WHERE pat_id = " & CStr(Pat_id) + " AND zeitpunkt > DokuDat - 550 ORDER BY zeitpunkt DESC", dbOpenDynaset) ' 1,5 Jahre
 rBr.Open "SELECT Quelldatum,Name,Zeitpunkt FROM `briefe` WHERE pat_id = " & CStr(aktDC.Pat_id) & " AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " AND name LIKE '%augen%' ORDER BY quelldatum DESC, zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
 If Not rBr.BOF Then
  AugUDat = rBr!Quelldatum
  If AugUDat = CDate(0) Then AugUDat = DatInStr(rBr!name, Year(rBr!Zeitpunkt))
  AugU = "Befund vo" + IIf(AugUDat = CDate(0), "n " + Format$(rBr!Zeitpunkt, "yyyy"), "m " + Format$(AugUDat, "dd/mm/yy")) + " vorliegend"
  If AugUDat = CDate(0) Then AugUDat = rBr!Zeitpunkt ' dann war die letzte Untersuchung vermutlich frühestens am 1.1. des Scan-Jahres
 Else
  Dim rAEin As New ADODB.Recordset
  rAEin.Open "SELECT ZeitPunkt FROM eintraege WHERE pat_id = " & CStr(aktDC.Pat_id) & " AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " AND art='aug' AND (NOT inhalt RLIKE 'nie|nicht|lange|länger|(kein|will|wird|macht).*Termin' OR inhalt RLIKE 'Veränd|in Ordnung') ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
  If Not rAEin.BOF Then
   AugUDat = rAEin!Zeitpunkt
  End If
 End If
' Dim rEin AS DAO.Recordset
 If lies.obMySQL Then
  sql = "SELECT * FROM `eintraege` WHERE pat_id = " & CStr(pid) & " AND ( art = 'aug' Or ((art in (" & latinis(artspezG) & ") AND ((inhalt like ""%augenb%"" AND not inhalt like ""%augenbl%"" AND not inhalt like ""%augen"") Or (inhalt like ""%augenarzt%"" Or inhalt like ""%augenärzt%"") Or (inhalt like ""% aa%"" AND not inhalt like ""% aag%""))) Or (art = ""aa"" Or art = ""augen"")))" & " AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " ORDER BY zeitpunkt DESC LIMIT 3"
 Else
  sql = "SELECT top 3 * FROM `eintraege` WHERE pat_id = " & CStr(pid) & " AND ( art = 'aug' Or ((art in (" & latinis(artspezG) & ") AND (inhalt like ""%augenb`!l`%"" Or inhalt like ""%augen`aä`rzt%"" Or inhalt like ""% aa`!g`%"")) Or (art = ""aa"" Or art = ""augen"")))" + " AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " ORDER BY zeitpunkt DESC"
 End If
' Set rEin = aktDCb.OpenRecordset(sql)
 Set raEi = Nothing
 raEi.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 Do While Not raEi.EOF
  AugU = AugU & IIf(LenB(AugU) <> 0, vbCrLf & vbTab, "") & "Eintrag am " & Format$(raEi!Zeitpunkt, "d/m/yy") & ":" + vbTab + raEi!Inhalt
  raEi.Move 1
 Loop
 Dim uebw As ADODB.Recordset 'DAO.Recordset
 Set uebw = New ADODB.Recordset '' 22.3.2012 wegen Möglichkeit Fehler
'  fiabfr = "SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, form_abk FROM (((`forminhfeld` LEFT JOIN `forminhkopf` ON `forminhfeld`.foid=`forminhkopf`.foid) LEFT JOIN `formulare` ON `formulare`.formid=`forminhkopf`.form_id) LEFT JOIN `forminhaltfeld` ON `forminhfeld`.feldvw=`forminhaltfeld`.feldvw) LEFT JOIN `forminhaltfeldinh` ON `forminhfeld`.feldinhvw=`forminhaltfeldinh`.feldinhvw "
'  Call uebw.Open(fiabfr & " WHERE pat_id = " & CStr(Pid) + " AND form_abk = ""uew"" AND feld = ""Ueberweisung_an"" AND feldinh = ""Augenheilkunde"" AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly)
' ohne Indexvorgabe Aufruf 0,4s durch Auswahl von Feldinh als Index, falls Bedingung feldinh = 'Augenheilkunde' enthalten
' Möglichkeit 1:
'  fiabfr = "SELECT * FROM (SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, form_abk FROM `forminhkopf` USE INDEX(auswahl) LEFT JOIN `formulare` ON `formulare`.formid = `forminhkopf`.form_id LEFT JOIN `forminhfeld` ON `forminhfeld`.foid = `forminhkopf`.foid LEFT JOIN `forminhaltfeld` ON `forminhfeld`.feldvw=`forminhaltfeld`.feldvw LEFT JOIN `forminhaltfeldinh` ON `forminhfeld`.feldinhvw=`forminhaltfeldinh`.feldinhvw WHERE pat_id = " & Pat_id & " AND form_abk = 'uew') AS i WHERE feldinh = 'Augenheilkunde' AND feld = 'Ueberweisung_an' AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " ORDER BY zeitpunkt DESC"
' Möglichkeit 2:
  fiabfr = "SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, form_abk FROM formulari WHERE pat_id=" & Pat_id & " AND form_abk='uew' AND feld ='Ueberweisung_an' AND feldinh ='Augenheilkunde' AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " ORDER BY zeitpunkt DESC"
  Call uebw.Open(fiabfr, DBCn, adOpenDynamic, adLockReadOnly)
' Set uebw = aktDCb.OpenRecordset(, dbOpenDynaset)
 If Not uebw.BOF Then
  If AugUDat = CDate(0) Or uebw!Zeitpunkt > AugUDat Then
   AugU = AugU + IIf(LenB(AugU) = 0, vNS, ", ") + "Untersuchung veranlasst am " + Format$(uebw!Zeitpunkt, "dd/mm/yy")
  End If
 End If
 If AugU = "" And tfeld(![Augensp zuletzt]) <> "" Then
  AugU = ![Augensp zuletzt]
  If tfeld(![Augensp Befund]) <> "" Then
   AugU = AugU + " (Befund: " + ![Augensp Befund] + ")"
  End If
 End If
 If AugUDat > DokuDat - 500 Then aktDC.aug = durchg Else aktDC.aug = veranl
 Dim crpos1%, crpos2%
 crpos1 = InStr(AugU, vbCrLf)
 crpos2 = InStr(crpos1 + 1, AugU, vbCrLf)
 If crpos1 <> 0 Then
  If crpos2 <> 0 Then
   AugU = Replace$(Left$(AugU, crpos2), vbTab, ", ")
  Else
   AugU = Left$(AugU, crpos1)
  End If
  AugU = Replace$(AugU, vbCr, "")
  AugU = Replace$(AugU, vbLf, "")
 End If
 If mitStr Then TabPr "Augenuntersuchung:", AugU
 
 If mitStr Then TabPr "Empf. Dok'intervall:", IIf(aktDC.hbEmpf = halten And (Not obPosi(tfeld(![Unterzucker pM])) Or InStrB(tfeld(![Unterzucker pM]), "selten") <> 0), "viertel- oder halbjährlich", "vierteljährlich")
End With ' raana

If mitStr Then DMPString = DmPStrS
lDMPPat_id = pid
üdt = aktDC
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If Err.Number = -2147467259 Then ' Server has gone away ' 30.8.17 ungültiger Eigenschaftswert führt zur Dauerschleife
'If Err.Number = -2147467259 Or Err.Number = -2147217887 Then ' Server has gone away / ungültiger Eigenschaftswert
 If DBCn.State <> 0 Then DBCn.Close
 Call DBCnOpen
 Resume nochmal
End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DMPString/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' DMPString

Function zuJahr$(ByVal Ursp$, Optional dokzt As Date)
 Dim buch$, pos&
 If Ursp = "?" Or Ursp = "" Then
  If dokzt = 0 Then dokzt = Now
  zuJahr = Year(dokzt)
 Else
  Ursp = Trim$(Replace$(Replace$(Replace$(Ursp, "ca.", ""), "ca", ""), "seit", ""))
  pos = InStr(Ursp, "/")
  If pos > 0 Then
   Ursp = Mid$(Ursp, pos + 1)
   buch = Mid$(Ursp, 3)
   If Not IsNumeric(buch) Then
    buch = Left(Ursp, 1)
    If Not IsNumeric(buch) Then zuJahr = vNS: Exit Function
    If buch > 5 Then
     Ursp = "19" & Ursp
    Else
     Ursp = "20" & Ursp
    End If
   End If
  End If
  pos = InStr(Ursp, " ")
  If pos > 0 Then
   If IsNumeric(Left(Ursp, 1)) Then
    Ursp = Left(Ursp, pos - 1)
   Else
    Ursp = Mid$(Ursp, pos + 1)
   End If
  End If
  pos = InStr(Ursp, "-")
  If pos > 0 Then
   Ursp = Left(Ursp, pos - 1)
  End If
  pos = InStr(Ursp, "a")
  If pos > 0 Then
   Ursp = Left(Ursp, pos - 1)
  End If
  pos = InStr(Ursp, "J")
  If pos > 0 Then
   Ursp = Left(Ursp, pos - 1)
  End If
  If InStrB(Ursp, ".") > 0 Then
   If IsDate(Ursp) Then
    zuJahr = Year(Ursp)
   End If
  Else
   If Not IsNumeric(Ursp) Then
    If dokzt = 0 Then dokzt = Now
    zuJahr = Year(dokzt)
   ElseIf Ursp < 50 Then
    If dokzt = 0 Then dokzt = Now
    zuJahr = Year(dokzt - Ursp * 365)
   ElseIf Ursp < 100 Then
    zuJahr = "19" & Ursp
   Else
    zuJahr = Ursp
   End If
  End If
 End If
End Function ' zuJahr

Function TherUmw$(therart As TherapieArt)
 Select Case therart
  Case Diät: TherUmw = "Diät?"
  Case offen: TherUmw = vNS
  Case ict: TherUmw = "ICT"
  Case csii: TherUmw = "CSII"
  Case ct: TherUmw = "CT"
  Case komb: TherUmw = "Komb"
  Case OAD: TherUmw = "OAD"
  Case glp1: TherUmw = "GLP1"
  Case glp1ins: TherUmw = "GLP1Ins"
  Case glp1ict: TherUmw = "GLP1ICT"
  Case Else
    MsgBox "Unvorhergesehene Therapieart in TherUmw: " & therart
    Tüt 1760, 1000
 End Select
End Function ' TherUmw

Function ImportFolderHerricht()
 Dim fld As Folder, Fil As File, FPath$
 On Error GoTo fehler
 Call VerzPrüf(hVerz)
 Set fld = FSO.GetFolder(hVerz)
 For Each Fil In fld.Files
  If Right$(Fil.name, 4) = ".BDT" Then
   FPath = Fil.path
   On Error Resume Next
   Name FPath As Replace$(FPath, ".BDT", ".txt")
   Kill FPath
   On Error GoTo fehler
  End If
 Next Fil
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ImportFolderHerricht/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ImportFolderHerricht()
'Function LeistungsExport0()
' Dim op$, BDT AS New BDTSchreib
' ON Error GoTo fehler
'' Call ImportFolderHerricht
'' LeistBDT = hVerz + "LEIST " + Format$(Now, "dd/mm/yy HH.MM") + ".BDT"
'' Open LeistBDT For Output AS #310
'  If Not BDT.Start(hVerz, "Leist") Then
'   Exit Function
'  End If
'  Print #310, "01380000020"
'  Print #310, "014810000082"
'  Print #310, "01691006419153"
'  Print #310, "017910309042005"
'  Print #310, "0129105001"
'  Print #310, "01091062"
'  Print #310, "01380000022"
'  Print #310, "014810000107"
'  Print #310, "014921001/99"
'  Print #310, "014921302/94"
'  Print #310, "01096001"
'  Print #310, "025960101011980" + Format$(Now, "ddmmyyyy")
'  Print #310, "017960214290200"
'  Print #310, "01380000010"
'  Print #310, "014810000315"
'  Print #310, "0160101A001011"
'  Print #310, "0260102TurboMed EDV GmbH"
'  Print #310, "0250103TurboMed@Windows"
'  Print #310, "0180104IBM PC/AT"
'  Print #310, "01602016419153"
'  Print #310, "01002021"
'  Print #310, "0220203Gerald Schade"
'  Print #310, "0500204FA Innere und Allgemeinmedizin (Hausarzt)"
'  Print #310, "0290205Mittermayerstraße 13"
'  Print #310, "014021585221"
'  Print #310, "0150216Dachau"
'  Print #310, "024020808131 / 616 380"
'  Print #310, "024020908131 / 616 381"
' Close #310
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 Then
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#End If
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LeistungsExport0/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End Function ' LeistungsExport0

' in doDiagnosenExport und Leistungsexport1
Function FallExport(BDT As BDTSchreib, Pat_id&, Datu As Date, ByRef Lanr&, Optional nurKasse%, Optional Arztnr&)
 Dim rFa As New ADODB.Recordset
 Dim op$
 On Error GoTo fehler
 rFa.Open "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_id & " AND bhfb<=" & DatFor_k(Datu) & " AND (bhfe1 >= " & DatFor_k(Datu) & " OR bhfe1=18991230) " & IIf(nurKasse, " AND schgr <> '90' ", "") & " ORDER BY bhfe1", DBCn, adOpenStatic, adLockReadOnly
 If rFa.BOF Then ' wenn kein Fall in die Zeitschiene paßt, dann den letzten vor dem Datum nehmen
  Set rFa = Nothing
  rFa.Open "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_id & " AND bhfb<=" & DatFor_k(Datu) & IIf(nurKasse, " AND schgr <> '90' ", "") & " ORDER BY bhfe1 DESC", DBCn, adOpenStatic, adLockReadOnly
  If rFa.BOF Then ' wenn kein Fall in die Zeitschiene paßt, dann den letzten nehmen
   Set rFa = Nothing
   rFa.Open "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_id & IIf(nurKasse, " AND schgr <> '90' ", "") & " ORDER BY bhfe1 DESC", DBCn, adOpenStatic, adLockReadOnly
   If rFa.BOF Then
    Exit Function
   End If
  End If ' rFa.BOF Then
 End If ' rFa.BOF Then
 If Arztnr <> 0 And rFa!lanrid <> Arztnr Then Exit Function
 FallExport = 1
 Lanr = nz(rFa!Lanr,0)
 Call BDT.Satzart(IIf(rFa!SchGr = "90", "0190", "0102")) ' 80000 Satzidentifikation
 Call BDT.PatID(rFa!Pat_id)
     BDT.SAdd "3635", "TM#" & rFa!Lanr
     BDT.SAdd "3636", "TM#" & "641915300"
     BDT.SAdd "4101", rFa!Quartal
     If rFa!SchGr <> "90" Then
      BDT.DAdd "4102", rFa!ausgst
      If rFa!VKNr <> vNS And Not IsNull(rFa!VKNr) Then
       BDT.SAdd "4104", rFa!VKNr
      End If
     End If
     If Not IsNull(rFa!KtrAbrB) And rFa!KtrAbrB <> vNS Then ' bei Privaten
      BDT.SAdd "4106", rFa!KtrAbrB
     End If
     If Not IsNull(rFa!AbrAr) And rFa!AbrAr <> vNS Then  ' bei Privaten
      BDT.SAdd "4107", rFa!AbrAr
     End If
     If rFa!SchGr <> "90" Then
      BDT.DAdd "4109", rFa!lVorl
      BDT.TAdd "4110", rFa!lVorl
      BDT.SAdd "4111", rFa!IK
      If Not IsNull(rFa!KVKs) And rFa!KVKs <> vNS Then ' bei Pat_id 43
       BDT.SAdd "4112", rFa!KVKs
      End If
      If Not IsNull(rFa!KVKserg) And rFa!KVKserg <> vNS Then
       BDT.SAdd "4113", rFa!KVKserg
      End If
     End If ' rFa!SchGr <> "90" Then
     If Not IsNull(rFa!GebOr) And rFa!GebOr <> vNS Then ' bei Privaten
      BDT.SAdd "4121", rFa!GebOr
     End If
     If Not IsNull(rFa!AbrGb) And rFa!AbrGb <> vNS Then ' bei Privaten
      BDT.SAdd "4122", rFa!AbrGb
     End If ' Not ISNULL(rFa!GebOr) Then ' bei Privaten
     BDT.SAdd "4144", "TM#" & nz(rFa!TMFNr,Space$(11))
     BDT.DAdd "4150", rFa!BhFB
     BDT.DAdd "4151", rFa!BhFE1
'     op = format$(3 + 4 + 8, "000") + "4152" + IIf(rFa!BhFE2 = 0, "00000000", format$(rFa!BhFE2, "ddmmyyyy"))
'     Print #310, ZSU(op)
     Dim Üw$
     BDT.SAdd "4218", "TM#" & rFa!ÜbWVLANR & "#" & rFa!ÜbWVBSNR & "#" & rFa!ÜbWVKVNR
     Üw = rFa!Übwr ' nz(rFa!AndÜw,vNS)
' 24.9.11: nicht nötig
'     If Üw <> vNS Then
'      BDT.SAdd "4219", Üw
'     End If
'     If Not ISNULL(rFa!ÜWZiel) AND rFa!ÜWZiel <> "" Then
'      BDT.SAdd "4220", rFa!ÜWZiel
'     End If
' 4285 Überweiser
     BDT.Add "4239" & rFa!SchGr
' 4401
' 4402
#If False Then
     If (Not IsNull(rFa!GOÄKatNr) And rFa!GOÄKatNr <> "") Or (Not IsNull(rFa!GOÄKatName) And rFa!GOÄKatName <> "") Then
      Üw = "TM#" + CStr(rFa!GOÄKatNr) + "#" + CStr(rFa!GOÄKatName)
      op = Format$(3 + 4 + Len(Üw), "000") + "4580" + Üw
      Print #310, ZSU(op)
     End If
#End If
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FallExport/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' FallExport

Function LeistungsExport1(BDT As BDTSchreib, Pat_id&, Leist$, Datu As Date, Optional QUZeit$, Optional nurKasse%, Optional Arztnr&)
 Dim rNa As New ADODB.Recordset, Lanr&
 rNa.Open "SELECT * FROM `namen` WHERE pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
 Lese.Ausgeb "Trage Leistung " & Leist & " für Pat_ID " & Pat_id & " (" & rNa!Nachname & ", " & rNa!Vorname & ") mit Datum " & Format$(Datu, "dd.mm.yyyy") & " ein.", True
 Open pVerz & LEDatei For Append As #347
 Print #347, Pat_id & ": " & rNa!Nachname & ", " & rNa!Vorname & ", " & Format(Datu, "dd.mm.yyyy")
 Close #347
 If FallExport(BDT, Pat_id, Datu, Lanr, nurKasse, Arztnr) <> 1 Then Exit Function
 LeistungsExport1 = 1
     BDT.DAdd "5000", Datu
     Dim UZeit$
     If Leist = "9995" Then
      UZeit = "0001"
     ElseIf IsNull(QUZeit) Then
      UZeit = "1900"
     ElseIf QUZeit = CDate(0) Then
      UZeit = "1900"
     Else
      UZeit = Format$(QUZeit, "hhmm")
     End If
     BDT.Add "6201" & UZeit
     BDT.SAdd "6203", "TM#?##" ' 17.4.11
     BDT.SAdd "3635", "TM#" & Lanr
     BDT.SAdd "3636", "TM#" & "641915300"
     BDT.SAdd "5001", Leist
     BDT.SAdd "5098", "641915300"
     BDT.SAdd "5099", CStr(Lanr)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LeistungsExport1/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LeistungsExport1

#If False Then
Function LeistungsExport1a(BDT As BDTSchreib, Pat_id&, Leist$, Datu As Date, Optional QUZeit$, Optional nurKasse%, Optional Arztnr%)
' Dim q AS DAO.Recordset
 Dim rFa As New ADODB.Recordset, rNa As New ADODB.Recordset
 Dim op$
 On Error GoTo fehler
' Set rFa = TabÖff("faelle", "Auswahl")
 rFa.Open "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_id & " AND bhfe1 >= " & DatFor_k(Datu) & IIf(nurKasse, " AND schgr <> '90' ", "") & " ORDER BY bhfe1", DBCn, adOpenStatic, adLockReadOnly
 If rFa.BOF Then ' wenn kein Fall in die Zeitschiene paßt, dann den letzten nehmen
  Set rFa = Nothing
  rFa.Open "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_id & IIf(nurKasse, " AND schgr <> '90' ", "") & " ORDER BY bhfe1 DESC", DBCn, adOpenStatic, adLockReadOnly
 End If
 If rFa.BOF Then Exit Function
 If Arztnr <> 0 And rFa!lanrid <> Arztnr Then Exit Function
 rNa.Open "SELECT * FROM `namen` WHERE pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
 Lese.Ausgeb "Trage Leistung " & Leist & " für Pat_ID " & Pat_id & " (" & rNa!Nachname & ", " & rNa!Vorname & ") mit Datum " & Format$(Datu, "dd.mm.yyyy") & " ein.", True
 LeistungsExport1a = 1
 Call BDT.Satzart(IIf(rFa!SchGr = "90", "0190", "0102")) ' 80000 Satzidentifikation
' bei 0101 entstehen bei zwei Aufrufen fehlerfrei zwei neue Kassenfaelle, jeder mit der Leistung
' bei 190 entsteht ein neuer Privatfall, bei 6100 läuft alles ohne Fehler durch, aber keine Leistung steht drin
' bei 6200 entsteht ein neuer Kassenfall
'     op = format$(3 + 4 + 4, "000") + "8000" + CStr(f!s8000)
'     Print #310, zsu(op)
'     op = format$(3 + 4 + 5, "000") + "8100" + CStr(f!s8100)
'     Print #310, zsu(op)
 Call BDT.PatID(rFa!Pat_id)
'#If False Then
'     If 1 = 0 Then 'Auswirkung bisher nicht geprüft 31.7.05 (3x)
'      op = Format$(3 + 4 + Len(n!NVorsatz), "000") + "3100" + CStr(n!NVorsatz)
'      Print #310, ZSU(op)
'     End If
'     op = Format$(3 + 4 + Len(n!Nachname), "000") + "3101" + CStr(n!Nachname)
'     Print #310, ZSU(op)
'     op = Format$(3 + 4 + Len(n!Vorname), "000") + "3102" + CStr(n!Vorname)
'     Print #310, ZSU(op)
'     op = Format$(3 + 4 + 8, "000") + "3103" + Format$(n!GebDat, "ddmmyyyy")
'     Print #310, ZSU(op)
'     op = Format$(3 + 4 + Len(nz(n!Versichertennummer,vNS)), "000") + "3105" + nz(n!Versichertennummer,vNS)
'     Print #310, ZSU(op)
''     op = format$(3 + 4 + Len(n!Straße), "000") + "3107" + n!Straße
''     Print #310, ZSU(op)
''     op = format$(3 + 4 + Len(n!Plz), "000") + "3112" + n!Plz
''     Print #310, ZSU(op)
''     op = format$(3 + 4 + Len(n!Ort), "000") + "3113" + n!Ort
''     Print #310, ZSU(op)
'''     GoTo leistungdirekt
'     If Not ISNULL(n!KVKStatus) AND n!KVKStatus > 0 AND Trim$(n!KVKStatus) <> "" Then
'      op = Format$(3 + 4 + Len(n!KVKStatus), "000") + "3108" + n!KVKStatus
'      Print #310, ZSU(op)
'     End If
'     op = Format$(3 + 4 + 1, "000") + "3110" + IIf(n!Geschlecht = "m", "1", IIf(n!Geschlecht = "w", "2", ""))
'     Print #310, ZSU(op)
'#End If
'#If False Then
'     If q!SchGr <> "90" Then
'      op = Format$(3 + 4 + 5, "000") + "4101" + Quartal
'      Print #310, ZSU(op)
'      op = Format$(3 + 4 + 8, "000") + "4102" + Format$(rFa!ausgst, "ddmmyyyy")
'      Print #310, ZSU(op)
'      op = Format$(3 + 4 + Len(rFa!VKNr), "000") + "4104" + rFa!VKNr
'      Print #310, ZSU(op)
'     End If
'#End If
     BDT.SAdd "3635", "TM#" & rFa!Lanr
     BDT.SAdd "3636", "TM#" & "641915300"
     BDT.SAdd "4101", rFa!Quartal
     If rFa!SchGr <> "90" Then
      If rFa!ausgst <> "1899-12-30 00:00:00" Then
       BDT.DAdd "4102", rFa!ausgst
      End If
      If rFa!VKNr <> vNS And Not IsNull(rFa!VKNr) Then
       BDT.SAdd "4104", rFa!VKNr
      End If
     End If
     If Not IsNull(rFa!KtrAbrB) And rFa!KtrAbrB <> vNS Then ' bei Privaten
      BDT.SAdd "4106", rFa!KtrAbrB
     End If
     If Not IsNull(rFa!AbrAr) And rFa!AbrAr <> vNS Then  ' bei Privaten
      BDT.SAdd "4107", rFa!AbrAr
     End If
     If rFa!SchGr <> "90" Then
      BDT.DAdd "4109", rFa!lVorl
      BDT.TAdd "4110", rFa!lVorl
      BDT.SAdd "4111", rFa!IK
      If Not IsNull(rFa!KVKs) And rFa!KVKs <> vNS Then ' bei Pat_id 43
       BDT.SAdd "4112", rFa!KVKs
      End If
      If Not IsNull(rFa!KVKserg) And rFa!KVKserg <> vNS Then
       BDT.SAdd "4113", rFa!KVKserg
      End If
     End If ' rFa!SchGr <> "90" Then
     If Not IsNull(rFa!GebOr) And rFa!GebOr <> vNS Then ' bei Privaten
      BDT.SAdd "4121", rFa!GebOr
     End If
     If Not IsNull(rFa!AbrGb) And rFa!AbrGb <> vNS Then ' bei Privaten
      BDT.SAdd "4122", rFa!AbrGb
     End If ' Not ISNULL(rFa!GebOr) Then ' bei Privaten
     BDT.SAdd "4144", "TM#" & nz(rFa!TMFNr,Space$(11))
     BDT.DAdd "4150", rFa!BhFB
     BDT.DAdd "4151", rFa!BhFE1
'     op = format$(3 + 4 + 8, "000") + "4152" + IIf(rFa!BhFE2 = 0, "00000000", format$(rFa!BhFE2, "ddmmyyyy"))
'     Print #310, ZSU(op)
     Dim Üw$
     BDT.SAdd "4218", "TM" & rFa!ÜbWVLANR & "#" & rFa!ÜbWVBSNR & "#" & rFa!ÜbWVKVNR
     Üw = rFa!Übwr ' nz(rFa!AndÜw,vNS)
     If Üw <> vNS Then
      BDT.SAdd "4219", Üw
     End If
     If Not IsNull(rFa!ÜWZiel) And rFa!ÜWZiel <> "" Then
      BDT.SAdd "4220", rFa!ÜWZiel
     End If
     BDT.Add "4239" & rFa!SchGr
#If False Then
     If (Not IsNull(rFa!GOÄKatNr) And rFa!GOÄKatNr <> "") Or (Not IsNull(rFa!GOÄKatName) And rFa!GOÄKatName <> "") Then
      Üw = "TM#" + CStr(rFa!GOÄKatNr) + "#" + CStr(rFa!GOÄKatName)
      op = Format$(3 + 4 + Len(Üw), "000") + "4580" + Üw
      Print #310, ZSU(op)
     End If
#End If
     BDT.DAdd "5000", Datu
     Dim UZeit$
     If Leist = "9995" Then
      UZeit = "0001"
     ElseIf IsNull(QUZeit) Then
      UZeit = "1900"
     ElseIf QUZeit = CDate(0) Then
      UZeit = "1900"
     Else
      UZeit = Format$(QUZeit, "hhmm")
     End If
     BDT.Add "6201" & UZeit
     BDT.SAdd "6203", "TM#?##" ' 17.4.11
     BDT.SAdd "3635", "TM#" & rFa!Lanr
     BDT.SAdd "3636", "TM#" & "641915300"
     BDT.SAdd "5001", Leist
 Close #310
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LeistungsExport1a/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LeistungsExport1a
#End If

Function dodoPorto()
 Open pVerz & LEDatei For Output As #347
 Print #347, "Porto vom " & Format(Now(), "dd.mm.yyyy HH:MM:SS:")
 Print #347, vbCrLf & "G.Schade:"
 Close #347
 dododoPorto 1
 Open pVerz & LEDatei For Append As #347
 Print #347, ""
 Print #347, "Dr.Kothny:"
 Close #347
 dododoPorto 2
 zeigan pVerz & LEDatei, vbNormalFocus
End Function 'dodoPorto

Function dododoPorto(Arztnr&)
 Dim Verz$
 Verz$ = pVerz & "gefaxt\"
 Dim BDT As New BDTSchreib
 Dim rs As New ADODB.Recordset, Zahl&, Zahl1&
 Dim DateiDatum As Date
 Dim Datei$, Pidpos&, pidp2%, pid&, Datum As Date, i&
#Const Manuell = True
#If Manuell Then
 Dim SL As New SortierListe, SSt As SortierString
 Set SL = Nothing
#End If
 On Error GoTo fehler
' If DBCn.State = 0 Then
'  Call DBVerb.cnVorb("", "anamnesebogen", "Patientendaten")
'  SetDBCn DBVerb.wCn
'  LVobMySQL = InStrB(UCase$(DBVerb.CnStr), "MYSQL") > 0
' End If
 If Not BDT.Start(hVerz, "Leist", Arztnr) Then
  Exit Function
 End If
 Call BDT.ImportFolderHerricht
 Call BDT.BDTKopf
' Call LeistungsExport0
 Datei = Dir(Verz)
 If LenB(Datei) = 0 Then
  MsgBox "Keine Faxe im Verzeichnis: '" & Verz & "'"
  Exit Function
 End If
 Do While LenB(Datei) <> 0
  If InStrB(Datei, " PID ") <> 0 And InStrB(LCase(Datei), ".pdf") <> 0 Then
   Dim geeignet%
   geeignet = geeignet + 1
   DateiDatum = FileDateTime(Verz & Datei)
   If DateiDatum >= QAnf(ZQuart(Now() - Verspätung)) And DateiDatum <= QEnd(ZQuart(Now() - Verspätung)) Then
    Pidpos = InStr(Datei, " PID ") + 5
    pidp2 = InStr(Pidpos, Datei, " ")
    pid = Mid$(Datei, Pidpos, pidp2 - Pidpos)
#If Manuell Then
    Set SSt = New SortierString
    SSt.Stri = pid & " " & Datum
    SL.sCAdd SSt, True
#End If
    Dim sql$
    If InStrB(Datei, " DMP-Daten ") <> 0 Then
     Pidpos = InStr(Datei, "DMP-Daten ") + 14 ' ... vom
     pidp2 = InStr(Pidpos, Datei, " ")
     Datum = Mid$(Datei, Pidpos, pidp2 - Pidpos)
'    Call LeistungsExport1(pid, "01600", datum, CDate("18:00"))
     sql = "SELECT * FROM `leistungen` WHERE pat_id = " & pid & " AND leistung = " & 40110 & " AND " & SelDatum("zeitpunkt", Datum) ' "date(zeitpunkt) = " & DatFor_k(datum)
     Set rs = Nothing
     rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
     If rs.BOF Then
      Zahl = Zahl + LeistungsExport1(BDT, pid, "40110", Datum, CDate("18:00"), True, Arztnr)
     End If
    ElseIf InStrB(Datei, "Arztbrief vom") <> 0 Then
     Pidpos = InStr(Datei, "Arztbrief vom ") + 14
     pidp2 = InStr(Pidpos, Datei, " ")
     Datum = Mid$(Datei, Pidpos, pidp2 - Pidpos)
'    Call LeistungsExport1(pid, "01601", datum, CDate("18:00"))
     sql = "SELECT * FROM leistungen WHERE pat_id = " & pid & " AND leistung = " & 40110 & " AND tag = " & DatFor_k(Datum)
     Set rs = Nothing
     rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
     If rs.BOF Then
      Zahl = Zahl + LeistungsExport1(BDT, pid, "40110", Datum, CDate("18:00"), True, Arztnr)
     End If
    End If ' InStrB(Datei, " DMP-Daten ") <> 0 Then elseif
   End If ' DateiDatum >= QAnf(ZQuart(Now() - Verspätung)) And DateiDatum <= QEnd(ZQuart(Now() - Verspätung)) Then
  End If ' InStrB(Datei, " PID ") <> 0 And InStrB(LCase(Datei), ".pdf") <> 0 Then
  Datei = Dir
 Loop ' While LenB(Datei) <> 0
 
' sql = "SELECT b.* FROM (SELECT pat_id, date(zeitpunkt) AS datum, time(zeitpunkt) AS zeit FROM `briefe` b WHERE (name like '%brief%' Or name like '%nachricht%') AND zeitpunkt BETWEEN " & lQAnfuEnd(Str(Verspätung)) & " GROUP BY pat_id, date(zeitpunkt)) AS b LEFT JOIN `leistungen` l ON b.pat_id = l.pat_id  AND leistung like '4012%' WHERE ISNULL(leistung)" ' DatFor_k(QAnf(ZQuart(Now() - Verspätung))) & " AND " & DatFor_k(QEnd(ZQuart(Now() - Verspätung)))
' ktag fehlerhaft
 sql = "SELECT b.* FROM (SELECT pat_id, date(zeitpunkt) Datum, TIME(zeitpunkt) zeit FROM `briefe` b WHERE (name LIKE '%brief%' Or name LIKE '%nachricht%') AND zeitpunkt BETWEEN  CONCAT(year(now()-interval  25 day),'-',(quarter(now()-interval  25 day)-1)*3+1,'-01') AND CONCAT(year(now()-interval  25 day)+ quarter(now()-interval  25 day) div 4 ,'-',((quarter(now()-interval  25 day)-1)*3+4) mod 12,'-01')-interval 0 day  GROUP BY pat_id, date(zeitpunkt)) AS b LEFT JOIN `leistungen` l ON b.pat_id = l.pat_id  AND DATE(l.zeitpunkt) = b.datum AND leistung LIKE '4012%' WHERE ISNULL(leistung)"
 Set rs = Nothing
 rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
 Zahl1 = 0
 Do While Not rs.EOF
  Zahl1 = Zahl1 + LeistungsExport1(BDT, rs!Pat_id, "40110", rs!Datum, CDate(rs!Zeit), True, Arztnr)
  rs.Move 1
 Loop
 BDT.Schreib
#If Manuell Then
    Dim PortoDat$
    PortoDat = pVerz & "\Porto " & Date & ".txt"
    On Error Resume Next
    Kill PortoDat
    On Error GoTo fehler
    Open PortoDat For Append As #303
    For i = 1 To SL.COUNT
     Print #303, SL.Item(i).Stri
    Next i
    Close #303
#End If
 Lese.Ausgeb "Datei '" & BDT.DMPImp & "' neu mit " & vbCrLf & Zahl & " Leistungen zu den Briefen in '" & Verz & "' und " & vbCrLf & Zahl1 & " Leistungen zu Arztbriefen in der Tabelle `briefe` für Kassenpatienten erstellt!, geeignet: " & geeignet, True
 Lese.Ausgeb sql, True
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doPorto/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dodoPorto

Function tuBriefeLeiDok(frm As Lese, Optional Arztnr&)
 On Error GoTo fehler
 Dim sql2$
 Dim rf As New ADODB.Recordset, rB As New ADODB.Recordset, rl As New ADODB.Recordset
 sql1 = "SELECT min(bhfb) AS bfb, min(schgr) AS schgr, pat_id, min(fid) AS fid FROM `faelle` WHERE schgr in ('21','23','24','00') AND quartal = """ & ZQuart(Now - Verspätung) & """ GROUP BY pat_id"
' sql1 = "SELECT f.fid AS fid, date(b.zeitpunkt) AS tag, b.* FROM `briefe` b LEFT JOIN `faelle` f ON b.pat_id = f.pat_id WHERE quartal = '" & ZQuart(Now - Verspätung) & "' AND schgr in ('21','23','24','00') AND b.zeitpunkt >= " & DatFor_k(QAnf(ZQuart(Now() - 20))) & " AND name like '%.doc' AND (name like '%brief%' Or name like '%dmp-daten%' Or name like '%nachricht an%') ORDER BY b.pat_id,b.zeitpunkt;"
' Call LeistungsExport0
 rf.Open sql1, DBCn, adOpenStatic, adLockReadOnly
 Do While Not rf.EOF
'  If rF!Pat_id = 2155 Then
  sql1 = "SELECT * FROM `briefe` WHERE pat_id = " & rf!Pat_id & " AND zeitpunkt >= " & DatFor_k(QAnf(ZQuart(Now() - 20))) & " AND name like '%.doc' AND (name like '%brief%' Or name like '%nachricht an%') AND zeitpunkt >= " & DatFor_k(#10/16/2007#)
  Set rB = Nothing
  rB.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
  If Not rB.EOF Then
   Set rl = Nothing
   rl.Open "SELECT * FROM `leistungen` WHERE fid = " & rf!FID & " AND leistung in ('01601')", DBCn, adOpenStatic, adLockReadOnly
   If rl.BOF Then
'    Call LeistungsExport1(rF!Pat_id, "01601", min(DateValue(rB!Zeitpunkt), #12/31/2007#), CDate("18:00"))
    Call LeistungsExport1(rf!Pat_id, "40110", Min(DateValue(rB!Zeitpunkt), #12/31/2007#), CDate("18:00"), , , Arztnr)
   End If
  Else
   sql1 = "SELECT * FROM `briefe` WHERE pat_id = " & rf!Pat_id & " AND zeitpunkt >= " & DatFor_k(QAnf(ZQuart(Now() - 20))) & " AND name like '%.doc' AND (name like '%dmp-daten%') AND zeitpunkt >= " & DatFor_k(#10/16/2007#)
   Set rB = Nothing
   rB.Open sql1, DBCn, adOpenStatic, adLockReadOnly
   If Not rB.EOF Then
    Set rl = Nothing
    rl.Open "SELECT * FROM `leistungen` WHERE fid = " & rf!FID & " AND leistung in ('01600')", DBCn, adOpenStatic, adLockReadOnly
    If rl.BOF Then
'     Call LeistungsExport1(rF!Pat_id, "01600", min(DateValue(rB!Zeitpunkt), #12/31/2007#), CDate("18:00"))
     Call LeistungsExport1(rf!Pat_id, "40110", Min(DateValue(rB!Zeitpunkt), #12/31/2007#), CDate("18:00"), , , Arztnr)
    End If
   End If
  End If
  
  rf.Move 1
 Loop
 MsgBox "Fertig meit Leistungsdoku"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in tuBriefeLeiDok/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' tuBriefeLeiDok

Function alleDMPLeiDok(frm As Lese, Optional Arztnr&)
 Dim DefDB$
 Dim rsa As New ADODB.Recordset, rMV As New ADODB.Recordset
 On Error GoTo fehler
' sql1 = "SELECT schgr, `faelle`.bhfb AS bfb, `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.vorname AS hvv, `hae`.nachname AS hvn, `hae`.fax1, hausaerzte.telefax, üwnnr FROM (((`faelle` inner join `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(mid$(`faelle`.üwnnr,1,2)¡""/""¡mid$(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & DefDB & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr WHERE quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr in (""24"",""23"",""21"",""00"") AND tkz = 0 ORDER BY nachname, vorname, schgr"
 sql1 = "SELECT bhfb AS bfb, schgr, pat_id FROM `faelle` WHERE schgr in (""24"",""23"",""21"",""00"") AND quartal = """ & ZQuart(Now - Verspätung) & """ "
' If lies.obMySQL Then
'  sql1 = replace$(sql1, "¡", ",")
' Else
'  sql1 = replace$(replace$(sql1, "concat", ""), "¡", " & ")
' End If
 rsa.Open "SELECT COUNT(DISTINCT pat_id) FROM (" & sql1 & ") AS innen", DBCn, adOpenStatic, adLockReadOnly
 frm.Ausgeb rsa.Fields(0) & " Fälle ", True
 Set rsa = Nothing
 rsa.Open sql1, DBCn, adOpenDynamic, adLockOptimistic
 Dim FZahl&, NZahl&
 Dim AltID&
 Dim anfang%, obzutr%
 Dim OkSql$
 Dim BDT As New BDTSchreib
 If Not BDT.Start(hVerz, "Leist", Arztnr) Then
  Exit Function
 End If
 Call BDT.BDTKopf
 Call BDT.ImportFolderHerricht
 Do While Not rsa.EOF
  obzutr = -1
  If rsa!SchGr = "00" Then ' bei eigenen nur die im DMP, da diese vermutlich noch einen anderen Hausarzt haben
    OkSql = "SELECT * FROM ((((`forminhkopf` LEFT JOIN `formulare` ON `forminhkopf`.form_id = `formulare`.formid) LEFT JOIN `forminhfeld` ON `forminhfeld`.foid = `forminhkopf`.foid) LEFT JOIN `forminhaltfeld` ON `forminhaltfeld`.feldvw = `forminhfeld`.feldvw) LEFT JOIN `forminhaltfeldinh` ON `forminhaltfeldinh`.feldinhvw = `forminhfeld`.feldinhvw) LEFT JOIN `namen` ON `forminhkopf`.pat_id = `namen`.pat_id WHERE form_abk like ""DMPDTYP" & "%" & """ AND feld like """ & "%" & "datum"" AND `namen`.Pat_id = " & rsa!Pat_id & ";"
    Set rMV = Nothing
    rMV.Open OkSql, DBCn, adOpenDynamic, adLockReadOnly
    If rMV.EOF Then obzutr = 0
  End If
  If obzutr Then
    Set rMV = Nothing
    rMV.Open "SELECT * FROM `leistungen` WHERE pat_id = " & rsa!Pat_id & " AND leistung = 1601 AND zeitpunkt >= " & DatFor_k(QAnf(ZQuart(Now - Verspätung))) & " AND zeitpunkt < " & DatFor_k(QEnd(ZQuart(Now - Verspätung))), DBCn, adOpenDynamic, adLockReadOnly
    If Not rMV.BOF Then obzutr = 0
  End If
  If obzutr Then
   If anfang Then
    If rsa!Pat_id <> AltID Then
' in MVZ nur die schicken, die nicht schon bei mir eingeschrieben
      Dim DaT As Date
      DaT = CDate("30.6.07")
      If rsa!bfb > DaT Then
       DaT = CDate("30.6.07")
      End If
'      Call LeistungsExport1(BDT, rsA!Pat_id, "01601", DaT, CDate("18:00"))
      Call LeistungsExport1(BDT, rsa!Pat_id, "40110", DaT, CDate("18:00"), , Arztnr)
      FZahl = FZahl + 1
    End If
    AltID = rsa!Pat_id
   End If
   anfang = True
  Else
   NZahl = NZahl + 1
  End If
'  endif
  rsa.Move 1
 Loop
 Call BDT.Schreib
 frm.Ausgeb "In " & FZahl & " Fällen Leistungen eingetragen, in " & NZahl & " weggelassen.", True
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in alleDMPLeiDok/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' alleDMPLeiDok

'Function alleDMPLeiDok1(frm AS Lese)
' Dim rsA AS New ADODB.Recordset, rMV AS New ADODB.Recordset
'
' sql1 = "SELECT `faelle`.bhfb AS bfb, `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.vorname AS hvv, `hae`.nachname AS hvn, `hae`.fax1, `hausaerzte`.telefax, üwnnr FROM (((`faelle` inner join `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(mid$(`faelle`.üwnnr,1,2)¡""/""¡mid$(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN `" & DefDB(DBCn) & "`.`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr WHERE quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr in (""24"",""23"",""21"") AND tkz = 0 AND ((icd like ""E11." & "%" & """ AND `hae`.DMPT2 <> 0) Or (icd like ""E10." & "%" & """ AND `hae`.DMPT1 <> 0)) AND not üwnnr in (""6493842"",""6496648"",""6419153"",""6491291"",""6488286"",""6420029"") ORDER BY nachname, vorname, schgr;"
''' Dr. colberg / Schmidt:
'' sql1 = "SELECT DISTINCT `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.vorname AS hvv, `hae`.nachname AS hvn, `hae`.fax1, `hausaerzte`.telefax, üwnnr FROM (((`faelle` inner join `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(mid$(`faelle`.üwnnr,1,2)¡""/""¡mid$(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & defdb & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr WHERE quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr in (""24"",""23"",""21"") AND tkz = 0 AND ((icd like ""E11." & "%" & """ AND `hae`.DMPT2 <> 0) Or (icd like ""E10." & "%" & """ AND `hae`.DMPT1 <> 0)) AND üwnnr in (""6419027"", ""6419421"",""6419418"") ORDER BY nachname, vorname, schgr;"
' ' Dr. Hofner 64/80054
'' sql1 = "SELECT DISTINCT `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.vorname AS hvv, `hae`.nachname AS hvn, `hae`.fax1, `hausaerzte`.telefax, üwnnr FROM (((`faelle` inner join `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(mid$(`faelle`.üwnnr,1,2)¡""/""¡mid$(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & defdb & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr WHERE quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr in (""24"",""23"",""21"") AND tkz = 0 AND ((icd like ""E11." & "%" & """ AND `hae`.DMPT2 <> 0) Or (icd like ""E10." & "%" & """ AND `hae`.DMPT1 <> 0)) AND üwnnr in (""6480054"") ORDER BY nachname, vorname, schgr;"
' If lies.obMySQL Then
'  sql1 = Replace$(sql1, "¡", ",")
' Else
'  sql1 = Replace$(Replace$(sql1, "concat", ""), "¡", " & ")
' End If
' rsA.Open sql1, DBCn, adOpenDynamic, adLockOptimistic
' Dim FZahl&, NZahl&
' Dim AltID&
' Dim anfang%
' Dim BDT AS New BDTSchreib
' If Not BDT.Start(hVerz, "Leist") Then
'  Exit Function
' End If
' Call BDT.ImportFolderHerricht
' Call BDT.BDTKopf
'' Call LeistungsExport0
' Do While Not rsA.EOF
'  If anfang Then
' ' If (ISNULL(rsA!telefax) AND ISNULL(rsA!fax1)) Or (Not ISNULL(rsA!telefax) AND Not ISNULL(rsA!fax1) AND rsA!telefax <> rsA!fax1) Then
'   If rsA!Pat_id <> AltID Then
'' in MVZ nur die schicken, die nicht schon bei mir eingeschrieben
'    Dim OkSql$
'    OkSql = "SELECT * FROM ((((`forminhkopf` LEFT JOIN `formulare` ON `forminhkopf`.form_id = `formulare`.formid) LEFT JOIN `forminhfeld` ON `forminhfeld`.foid = `forminhkopf`.foid) LEFT JOIN `forminhaltfeld` ON `forminhaltfeld`.feldvw = `forminhfeld`.feldvw) LEFT JOIN `forminhaltfeldinh` ON `forminhaltfeldinh`.feldinhvw = `forminhfeld`.feldinhvw) LEFT JOIN `namen` ON `forminhkopf`.pat_id = `namen`.pat_id WHERE form_abk like ""DMPDTYP" & "%" & """ AND feld like """ & "%" & "datum"" AND `namen`.Pat_id = " & rsA!Pat_id & ";"
'    Set rMV = Nothing
'    rMV.Open OkSql, DBCn, adOpenDynamic, adLockReadOnly
'    If rMV.EOF Or rsA!ÜWNNr <> "6419416" Then ' beim MVZ nicht die schon bei uns eingetragenen
'     Set rMV = Nothing
'     rMV.Open "SELECT * FROM `leistungen` WHERE leistung = 1600 AND zeitpunkt >= " & DatFor_k(CDate("28.6.07")) & " AND zeitpunkt < " & DatFor_k(CDate("28.6.07") + 1) & " AND pat_id = " & rsA!Pat_id, DBCn, adOpenDynamic, adLockReadOnly
'     If rMV.BOF Then
'      Dim DaT AS Date
'      DaT = CDate("16.6.07")
'      If rsA!bfb > DaT Then
'       DaT = CDate("30.6.07")
'      End If
''      Call LeistungsExport1(BDT, rsA!Pat_id&, "01600", DaT, CDate("13:00"))
'      Call LeistungsExport1(BDT, rsA!Pat_id&, "40120", DaT, CDate("13:00"))
'      FZahl = FZahl + 1
'     Else
'      NZahl = NZahl + 1
'     End If
'    End If
'    AltID = rsA!Pat_id
'   End If
'  End If
''  If rsA!Pat_id = 977 Then
'   anfang = True
''  endif
'  rsA.Move 1
' Loop
' frm.Ausgeb "In " & FZahl & " Fällen Leistungen eingetragen, in " & NZahl & " weggelassen.", True
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 Then
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#End If
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in alleDMPLeiDok1/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End Function ' alleDMPLeiDok1
Function MAX(D1, D2)
 If D1 > D2 Then MAX = D1 Else MAX = D2
End Function ' max

Function Min(D1, D2)
 If D1 < D2 Then Min = D1 Else Min = D2
End Function ' min

Sub tob()
 Dim dmpk&, dmpkhkk&, dmpcopdk&, dmpabk&, HzV&, DS&, hzvab As Date, dsab As Date
 Dim dmpb As Date, dmpkhkb As Date, dmpcopdb As Date, dmpabb As Date
 Call obhierdmpfn("DMP COPD MVZ 12.10.19", dmpk, dmpb, dmpkhkk, dmpkhkb, dmpcopdk, dmpcopdb, dmpabk, dmpabb, HzV, hzvab, DS, dsab)
 Debug.Print dmpk, dmpb, dmpcopdk, dmpcopdb, HzV, hzvab, DS, dsab
End Sub ' tob

Sub obhierdmpfn(Notiz$, dmpklass&, dmpbeg As Date, Optional dmpkhkklass&, Optional dmpkhkbeg As Date, Optional dmpcopdklass&, Optional dmpcopdbeg As Date, Optional dmpabklass&, Optional dmpabbeg As Date, Optional HzV&, Optional HzVbeg As Date, Optional DS&, Optional DSbeg As Date)
 Dim buch$, pos%, testd As Date, i%, notdat$(), ndRichtig%() ' letztes wahr, wenn notdat-Zeile für DMP verwertbar oder irrelevant
 Dim maxdHier(3) As Date, maxdHA(3) As Date, maxdNein(3) As Date, maxdAus(3) As Date
 Dim dmpArt% ' 0 = Diabetes, 1 = KHK, 2 = COPD, 3 = Asthma bronchiale
 On Error GoTo fehler
    If LenB(Notiz) <> 0 Then
     ReDim notdat(0)
     Do
      pos = InStr(Notiz, vbCrLf)
      If pos <> 0 Then
       notdat(UBound(notdat)) = UCase$(Trim$(Left(Notiz, pos - 1)))
       Notiz = Mid$(Notiz, pos + 2)
       ReDim Preserve notdat(UBound(notdat) + 1)
      Else
       notdat(UBound(notdat)) = UCase$(Trim$(Notiz))
       Notiz = vNS
       Exit Do
      End If
     Loop
     ReDim ndRichtig(UBound(notdat))
     Dim umw$
     For i = 0 To UBound(notdat)
      umw = vNS
      testd = 0
      For pos = Len(notdat(i)) To 1 Step -1
       buch = Mid$(notdat(i), pos, 1)
       Select Case buch
        Case " "
         If IsDate(umw) Then '(Mid$(notdat(i), pos + 1))
          testd = CDate(umw) '(Mid$(notdat(i), pos + 1))
          ndRichtig(i) = True ' b1) Datum als solches erkennbar
         End If
         If LenB(umw) <> 0 Then Exit For
        Case ".", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
         umw = buch & umw
        Case Else
         umw = vNS
       End Select
      Next pos
      
      If InStrB(notdat(i), "HZV") = 1 Or InStrB(notdat(i), "HM") = 1 Then
       HzV = 1
       HzVbeg = testd
      ElseIf InStrB(notdat(i), "DS") = 1 Then
       DS = 1
       DSbeg = testd
'      If InStrB(notdat(i), "DMP") = 0 Then
'       ndRichtig(i) = True ' a) für DMP irrelevant
      ElseIf InStrB(notdat(i), "DMP") = 1 Then
       dmpArt = 0
       If InStrB(notdat(i), "KHK") <> 0 Then dmpArt = 1 Else If InStrB(notdat(i), "COPD") <> 0 Then dmpArt = 2 Else If (InStrB(notdat(i), "ASTHMA") <> 0 Or InStrB(notdat(i), "AB") <> 0) Then dmpArt = 3
       If testd <> 0 Then
        If InStrB(notdat(i), "HIER") <> 0 Or InStrB(notdat(i), "BEI UNS") <> 0 Then
         maxdHier(dmpArt) = MAX(maxdHier(dmpArt), testd)
        ElseIf InStrB(notdat(i), " HA") <> 0 Or InStrB(notdat(i), " HÄ") <> 0 Then
         maxdHA(dmpArt) = MAX(maxdHA(dmpArt), testd)
        ElseIf InStrB(notdat(i), "NEIN") <> 0 Then
         maxdNein(dmpArt) = MAX(maxdNein(dmpArt), testd)
        ElseIf InStrB(notdat(i), "AUSGESCHRIEBEN") <> 0 Then
         maxdAus(dmpArt) = MAX(maxdAus(dmpArt), testd)
        Else ' z.B. DMP MVZ
         maxdHA(dmpArt) = MAX(maxdHA(dmpArt), testd)
         ndRichtig(i) = False ' b2) und "hier", " HA" oder "nein" enthalten
        End If ' instrb(notdat(i), ...
       End If ' test<>0
       If InStrB(notdat(i), "NEIN") <> 0 Then
        maxdNein(dmpArt) = Now() - 365
       ElseIf InStrB(notdat(i), "AUSGESCHRIEBEN") <> 0 Then
        maxdAus(dmpArt) = Now() - 365
       End If ' InStrB(notdat(i), "NEIN") <> 0 Then
      End If ' instrb(notdat(i), ...
     Next i
     For i = UBound(notdat) To 0 Step -1
      If ndRichtig(i) = 0 Then
       Notiz = notdat(i) & vbCrLf & Notiz
      End If
     Next i
     For dmpArt = 3 To 0 Step -1
        dmpklass = 0
        If maxdAus(dmpArt) > maxdHier(dmpArt) And maxdAus(dmpArt) > maxdHA(dmpArt) And maxdAus(dmpArt) > maxdHA(dmpArt) And maxdAus(dmpArt) <> 0 Then
         dmpklass = 4
         dmpbeg = maxdAus(dmpArt)
        ElseIf maxdHier(dmpArt) > maxdHA(dmpArt) And maxdHier(dmpArt) >= maxdNein(dmpArt) And maxdHier(dmpArt) <> 0 Then
          dmpklass = 3
          dmpbeg = maxdHier(dmpArt)
        ElseIf maxdNein(dmpArt) > maxdHA(dmpArt) And maxdNein(dmpArt) <> 0 Then
          dmpklass = 1
          dmpbeg = maxdNein(dmpArt)
        ElseIf maxdHA(dmpArt) <> 0 Then
          dmpklass = 2
          dmpbeg = maxdHA(dmpArt)
        End If
        If dmpklass <> 0 Then
         If dmpArt = 1 Then
          dmpkhkklass = dmpklass
          dmpkhkbeg = dmpbeg
          dmpklass = 0
          dmpbeg = 0
         ElseIf dmpArt = 2 Then
          dmpcopdklass = dmpklass
          dmpcopdbeg = dmpbeg
          dmpklass = 0
          dmpbeg = 0
         ElseIf dmpArt = 3 Then
          dmpabklass = dmpklass
          dmpabbeg = dmpbeg
          dmpklass = 0
          dmpbeg = 0
         End If
        End If
     Next dmpArt
    End If ' LenB(Notiz) <> 0 Then
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in obhierdmpfn/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub      ' obhierdmpfn

Function doVerdächtigeÜberweiser()
 Dim sql$, rf As New ADODB.Recordset, infos$(), taskid&, HAi%, AktQ$, obverd%
 Dim Datei$
 Datei$ = pVerz & "VerdächtigeÜberweiser.txt"
 AktQ = Left$(ZQuart(Now() - 14), 1) & Right$(ZQuart(Now() - 14), 2) ' 109
 sql = "SELECT v.pat_id,n.* FROM `aktfvs` v LEFT JOIN `namen` n ON v.pat_id = n.pat_id"
 Open Datei For Output As #355
 rf.Open sql, DBCn
 Do While Not rf.EOF
'   Call getHausarztAlt(Pid:=rF!Pat_id, Infos:=Infos(), obHAPrio:=False)
   Dim rFa() As Faelle
   Dim rKv1() As kvnrue
   getHausarzt1 infos, rFa, rKv1, False, rf!Pat_id
   HAi = 0
   obverd = False
   Do
    If InStrB(infos(10, HAi), "HA") <> 0 Then Exit Do
    HAi = HAi + 1
    If HAi > UBound(infos, 2) Then obverd = True: Exit Do
   Loop
   If Not obverd Then If HAi <> 0 And InStrB(infos(10, 0), AktQ) <> 0 And infos(1, 0) <> infos(1, HAi) Then obverd = True
   If obverd Then
    Print #355, rf!Pat_id & ": " & rf!Nachname & ", " & rf!Vorname & ":"
    Print #355, "  " & Left$(infos(10, 0) & Space$(10), 10) & ": " & infos(1, 0)
    If Not HAi > UBound(infos, 2) Then Print #355, "  " & Left$(infos(10, HAi) & Space$(10), 10) & ": " & infos(1, HAi)
    Ausgeb "verdächtig: " & rf!Pat_id & ": " & rf!Nachname & ", " & rf!Vorname
   Else
    Ausgeb "in Ordnung: " & rf!Pat_id & ": " & rf!Nachname & ", " & rf!Vorname
   End If
'   Stop
   rf.Move 1
 Loop
 Close #355
 taskid = zeigan(Datei, vbNormalFocus)
End Function
' 5.7.10: jetzt nicht mehr nötig, in PatListe integriert
'Function alleDMPs(frm AS Lese)
' Dim dszahl&
' Dim hae AS New adodb.Recordset
' Dim rsA AS New adodb.Recordset, rMV AS New adodb.Recordset
' Dim sql$, sql1$
' Dim AltID&
' Dim antw&
' antw = MsgBox("Wollen Sie wirklich alle DMP-Daten versenden?", vbYesNo)
' If antw = vbNo Then Exit Function
' QMdbAkt = QmdB ' Imortiere.bas uverz & "anamnese\quelle.mdb"
''SELECT `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.telefax AS hfax, `hausaerzte`.telefax FROM (((`faelle` inner join `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(mid$(`faelle`.üwnnr,1,2)¡"/"¡mid$(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & defdb & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr WHERE quartal = "42006" AND schgr in ("24","23","21") AND icd like "E11.%" AND tkz = 0 AND `hae`.DMPT2 <> 0 AND not üwnnr in ("6493842","6496648","6419153","6491291","6488286","6420029","6494531") ORDER BY schgr;
''Ausnahmen: Lindenthal, Pfeuffer, Rembold, Schneider, Stolzki, Wengeler
'
''  LEFT JOIN `kvaerzte`.`hae` ON CONCAT(mid$(`faelle`.üwnnr,1,2)¡""/""¡mid$(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & DefDB & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr) LEFT JOIN " & DefDB & ".`liuez` AS haxls ON `faelle`.üwnnr = haxls.kvnr
'
' sql1 = "SELECT `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, üwnnr, `namen`.notiz, `namen`.obhierdmp, icd FROM (((`" & DefDB(DBCn) & "`.`faelle` inner join `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `namen` ON `faelle`.pat_id = `namen`.pat_id) WHERE (quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr in (""00"",""24"") AND tkz = 0 AND icd REGEXP '^E1[0-4]\.' AND obhierdmp=0) ORDER BY nachname, vorname, schgr"
''' Dr. colberg / Schmidt:
''  sql1 = "SELECT DISTINCT `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.vorname AS hvv, `hae`.nachname AS hvn, `hae`.fax1, `hausaerzte`.telefax, haxls.fax, üwnnr, `namen`.notiz FROM (((((`faelle` inner join `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(mid$(`faelle`.üwnnr,1,2)¡""/""¡mid$(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & defdb & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr) LEFT JOIN " & defdb & ".`liuez` AS haxls ON `faelle`.üwnnr = haxls.kvnr) LEFT JOIN `namen` ON `faelle`.pat_id = `namen`.pat_id WHERE quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr in (""24"",""23"",""21"") AND tkz = 0 AND ((icd like ""E11." & "%" & """ AND `hae`.DMPT2 <> 0) Or (icd like ""E10." & "%" & """ AND `hae`.DMPT1 <> 0)) AND üwnnr in (""6419027"", ""6419421"",""6419418"") ORDER BY nachname, vorname, schgr"
' ' Dr. Hofner 64/80054
'' sql1 = "SELECT DISTINCT `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.vorname AS hvv, `hae`.nachname AS hvn, `hae`.fax1, `hausaerzte`.telefax, haxls.fax, üwnnr, `namen`.notiz FROM (((((`faelle` inner join `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(mid$(`faelle`.üwnnr,1,2)¡""/""¡mid$(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & defdb & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr) LEFT JOIN " & defdb & ".`liuez` AS haxls ON `faelle`.üwnnr = haxls.kvnr) LEFT JOIN `namen` ON `faelle`.pat_id = `namen`.pat_id WHERE quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr in (""24"",""23"",""21"") AND tkz = 0 AND ((icd like ""E11." & "%" & """ AND `hae`.DMPT2 <> 0) Or (icd like ""E10." & "%" & """ AND `hae`.DMPT1 <> 0)) AND üwnnr in (""6480054"") ORDER BY nachname, vorname, schgr"
' If lies.obMySQL Then
'  sql1 = Replace$(sql1, "¡", ",")
' Else
'  sql1 = Replace$(Replace$(sql1, "concat", ""), "¡", " & ")
' End If
' DBCn.Execute "DROP TABLE dmpausw"
' DBCn.Execute "CREATE TABLE dmpausw(pat_id int(10) unique key,name varchar(100), üwnnr varchar(10), icd varchar(10),fax varchar(20), adressat varchar(150)) comment 'Auswahl für alleDMPs'"
' DBCn.Execute "DELETE FROM dmpausw"
' sql1 = "SELECT * FROM (" & sql1 & ") AS innen GROUP BY pat_id" ' WHERE pat_id = 2193
' rsA.Open sql1, DBCn, adOpenDynamic, adLockOptimistic
'' For runde = 1 To 2
' Do While Not rsA.EOF
'   If rsA!Pat_id <> AltID Then
'    If True Or rsA!Pat_id = 50 Then
' '    If rsA!obhierdmp <> 0 Then ' falls nicht bei uns unterschrieben
'      lies.Ausgeb rsA!Pat_id, False, True
'      Call einDMP(rsA!Pat_id, rsA!ICD, dszahl, rsA!Nachname, rsA!Vorname, rsA!ÜWNNr, frm)
''     End If
'    End If
'   End If
'   AltID = rsA!Pat_id
'  rsA.Move 1
' Loop
' rsA.MoveFirst
' frm.GesBytes = dszahl
' dszahl = 0
'' Next runde
' lies.Ausgeb "Fertig mit alle DMPs!", True, True
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 Then
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#End If
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in alleDMPs/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End Function ' allDMPs

Function einDMP(Pat_id&, Optional ICD$, Optional dszahl&, Optional Nachname$, Optional Vorname$, Optional ÜWNNr$, Optional frm As Lese)
 Dim FaxNr$, infos$() ' Frau/Herrn, Vorn+Nachn, Straße, PLZ+Ort, Faxnr, S.g./Liebe, DMPTyp2, DMPTyp1
 Dim aktPatGefaxt$()
 Dim i%, j&, obdoppelt%, rAF&
 Dim docName$, Adressat$, fax1$, anfang%, runde%
 Const maxverschieden% = 1 ' maximale Zahl verschiedener Ärzte, an die Infos geschickt werden
    ReDim aktPatGefaxt(0)
     Call getHausarzt(pid:=Pat_id, infos:=infos()) ' , obHAPrio:=True)
'     Dim rs AS New adodb.Recordset
'     rs.Open "SELECT getha FROM `namen` WHERE pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
'     If rs.EOF Then Stop
'     If rs!getHA <> Infos(12, 0) Then Stop
'     Exit Function
     If LenB(ICD) = 0 Then ICD = IIf(infos(6, i) = "X", "E11", "E10")
     For i = 0 To Min(maxverschieden - 1, UBound(infos, 2))
'      If ((ICD Like "E11*" AND infos(6, i) = "X") Or (ICD Like "E10*" AND infos(7, i) = "X")) Then
      If ((ICD Like "E11*") Or (ICD Like "E10*")) Then
      ' 5.7.10: Diese beiden Zeilen wären notwendig für die zwischenzeitliche Umfunktionierung dieser Funktion,
      ' jetzt in Tabellen `namen` (obhierdmp, obdmp, getha0, getha1, getha2) und hareal
'       InsKorr DBCn, dbcns, "INSERT INTO `dmpausw` VALUES(" & Pat_id & ",'" & UmwfSQL(Nachname & ", " & Vorname) & "','" & Infos(12, i) & "','" & ICD & "','" & Infos(4, i) & "','" & UmwfSQL(Infos(14, i) & ", " & Infos(9, i) & ", " & Infos(3, i)) & "')", rAF
'       Exit Function
       Select Case infos(12, i)
        Case "6492325" ' Reitmeier, will keine DMP-Auswertungen mehr (23.12.09)
        Case "6419153" ' Schade
'        Case "6493950", "6419568" ' Schorten, Sewering
'        Case "6495818", "6491291" ' Pfeuffer, Krombholz
        Case Else
         fax1 = infos(4, i)
         If fax1 = "08131 25169" Then fax1 = "08131 273373" ' Dr. Stolzki falsche Faxnummer
         obdoppelt = False
         For j = 0 To UBound(aktPatGefaxt) - 1
          If aktPatGefaxt(j) = fax1 Then obdoppelt = True
         Next j
         If Not obdoppelt Then
          Debug.Print "Erstelle: ", Pat_id, Nachname, Vorname, fax1, ÜWNNr
'          If runde = 1 Then
'          Else
'           If pat_id = 1974 Then
           anfang = True
'           End If
           If anfang Then
            Call Ausgeb(Pat_id & ": " & Nachname & ", " & Vorname & ", " & fax1 & ", " & ÜWNNr)
            DoEvents
            Adressat = infos(10, i) & " " & infos(1, i)
            docName = do_DMPAusgebStandAlone(Pat_id, fax1, Adressat)
            If False Then
             Call FaxSend(docName, Adressat, fax1)
            End If
           End If ' 1 = 0
           aktPatGefaxt(UBound(aktPatGefaxt)) = fax1
           ReDim Preserve aktPatGefaxt(UBound(aktPatGefaxt) + 1)
'         Exit For ' Alternative zu obdoppelt
'          End If
          On Error Resume Next
          frm.Bytes = dszahl
          On Error GoTo fehler
          dszahl = dszahl + 1
         End If ' runde = 1
       End Select
      End If
     Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in einDMPs/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' einDMP

Function Ausgeb(AGStr$)
 Dim runde%, aktdat$
 aktdat = AutoBriefProtok
 On Error Resume Next
 For runde = 0 To 10
  Open aktdat For Append As #334 ' pverz & "zufaxen.txt" = AutoBriefProtok
  If Err.Number = 0 Then Exit For
  aktdat = Left$(aktdat, Len(aktdat) - 4) & "_" & Right$(aktdat, 4)
  Err.Clear
 Next runde
 If runde < 11 Then
  On Error GoTo fehler
  Print #334, AGStr
  Close #334
 End If
 Lese.Ausgeb AGStr, True
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Ausgeb/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
'' SpMinÜ und SpMaxÜ = Arrays mit minimalen / maximalen Überschriftenlängen
'Public Function TabAusgeb(rEinl AS ADODB.Recordset, Optional obMitausgeb% = False, Optional nz$ = vbCrLf, Optional ohneKopfz% = False, Optional SpMinÜ, Optional SpMaxÜ, Optional mitLeerZeilen% = False, Optional Datei$) AS CString
' Dim i%, maxL%(), Zrm%(), notNum%(), F1alt
' Const ZrmVorgabe% = 2
' ON Error GoTo fehler
' Set TabAusgeb = New CString
' ReDim maxL(rEinl.Fields.Count)
' ReDim Zrm(rEinl.Fields.Count)
' ReDim notNum(rEinl.Fields.Count)
' For i = 0 To UBound(Zrm)
'  Zrm(i) = ZrmVorgabe
' Next i
' Do While Not rEinl.EOF
'  For i = 0 To rEinl.Fields.Count - 1
'   If Len(rEinl.Fields(i).Value) > maxL(i) Then maxL(i) = Len(rEinl.Fields(i).Value)
'   If Not ISNULL(rEinl.Fields(i).Value) Then If Not IsNumeric(rEinl.Fields(i).Value) Then notNum(i) = True
'  Next i
'  rEinl.Move 1
' Loop
' If ohneKopfz = 0 Then
'  For i = 0 To rEinl.Fields.Count - 1
'   If Len(rEinl.Fields(i).Name) > maxL(i) Then maxL(i) = Len(rEinl.Fields(i).Name): Zrm(i) = ZrmVorgabe - 1
'  Next i
' End If
' If Not IsMissing(SpMinÜ) Then
'  For i = 0 To rEinl.Fields.Count - 1
'   If UBound(SpMinÜ) >= i Then
'    If maxL(i) < SpMinÜ(i) AND SpMinÜ(i) <> 0 Then
'     maxL(i) = SpMinÜ(i)
'    End If
'   End If
'  Next i
' End If
' If Not IsMissing(SpMaxÜ) Then
'  For i = 0 To rEinl.Fields.Count - 1
'   If UBound(SpMaxÜ) >= i Then
'    If maxL(i) > SpMaxÜ(i) AND SpMaxÜ(i) <> 0 Then
'     maxL(i) = SpMaxÜ(i)
'    End If
'   End If
'  Next i
' End If
' TabAusgeb.Clear
' If ohneKopfZ = 0 Then
'  For i = 0 To rEinl.Fields.Count - 1
'   TabAusgeb.Append Left$(rEinl.Fields(i).Name & Space$(maxL(i) + Zrm(i)), maxL(i) + Zrm(i))
'  Next i
'  TabAusgeb.Append nz
' End If ' not ohneKopfZ
' If Not rEinl.BOF Then
'  rEinl.MoveFirst
'  If mitLeerZeilen Then F1alt = rEinl.Fields(0)
'  Do While Not rEinl.EOF
'   If mitLeerZeilen Then If rEinl.Fields(0) <> F1alt Then TabAusgeb.Append nz: F1alt = rEinl.Fields(0)
'   For i = 0 To rEinl.Fields.Count - 1
'    If notNum(i) Then
'     TabAusgeb.Append Left$(rEinl.Fields(i).Value & Space$(maxL(i) + Zrm(i)), maxL(i) + Zrm(i))
'    Else
'     TabAusgeb.Append Right$(Space$(maxL(i)) & rEinl.Fields(i).Value, maxL(i)) & Space$(Zrm(i))
'    End If
'   Next i
'   TabAusgeb.Append nz
'   rEinl.Move 1
'  Loop
'  rEinl.MoveFirst
' End If ' not rEinl.BOF
' If obMitausgeb Then Lese.Ausgeb TabAusgeb.Value, True
' If LenB(Datei) <> 0 Then
'  If InStrB(Datei, ".") = 0 Then
'   Datei = pVerz & Datei & " " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".csv"
'  End If
'  Open Datei For Output AS #317
'  Print #317, TabAusgeb.Value
'  Close #317
'  zeigan Datei
' End If
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 Then
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#End If
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + nz + "LastDLLError: " + CStr(Err.LastDllError) + nz + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + nz + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in TabAusgeb/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End Function ' TabAusgeb

Sub dodoFollowUp(frm As Lese)
 Dim infos$()
 Dim sql$, sql1$
 Dim rP As New ADODB.Recordset
 Dim rB As New ADODB.Recordset
 Dim Zp As Date, Quartal$
 On Error GoTo fehler
 Zp = Now - 30
 Quartal = ZQuart(Zp)
' ab = InputBox("Ab welchem Pat. anfangen (0 = von vorne)?", "Rückfrage", 0)
 Call rP.Open("SELECT DISTINCT f.pat_id AS pat_id FROM `faelle` f LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id WHERE not schgr in ('41','42','43')  AND tkz = 0 ORDER BY f.pat_id DESC;", DBCn, adOpenDynamic, adLockReadOnly)
 Do While Not rP.EOF
  Call getHausarzt(rP!Pat_id, infos())
  If infos(1, 0) <> vNS And Not infos(1, 0) Like "*Schade" Then
   Set rB = Nothing
   sql = "SELECT * FROM `briefe` WHERE pat_id = " & rP!Pat_id & " AND ((name Like '%Brief an %Dr%' Or name Like '%Arztbrief%' Or name Like 'Brief an HA%' Or name Like 'Brief an HAe%' Or name like 'Brief an %') AND name NOT LIKE '%Entwurf%') ORDER BY zeitpunkt DESC"
   Call rB.Open(sql, DBCn, adOpenDynamic, adLockReadOnly)
   If Not rB.BOF Then
    If DateValue(lebe(rP!Pat_id)) - rB!Zeitpunkt > 365 Then
     Call tubriefStandalone(rP!Pat_id, True)
    End If
   End If
  End If
  rP.Move 1
 Loop
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doRestlicheBriefe/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' dodoFollowUp(frm AS Lese)

Sub doUngeschriebeneBriefe(frm As Lese)
 Dim infos$()
 Dim sql$, sql1$
 Dim rP As New ADODB.Recordset
 Dim rB As New ADODB.Recordset
 Dim Zp As Date, Quartal$
 On Error GoTo fehler
 Zp = Now - 30
 Quartal = ZQuart(Zp)
' ab = InputBox("Ab welchem Pat. anfangen (0 = von vorne)?", "Rückfrage", 0)
 Call rP.Open("SELECT DISTINCT f.pat_id AS pat_id FROM `faelle` f LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id WHERE not schgr in ('41','42','43')  AND tkz = 0  ORDER BY f.pat_id DESC;", DBCn, adOpenDynamic, adLockReadOnly)
 Do While Not rP.EOF
  Call getHausarzt(rP!Pat_id, infos())
  If infos(1, 0) <> vNS And Not infos(1, 0) Like "*Schade" Then
   Set rB = Nothing
   sql = "SELECT * FROM `briefe` WHERE pat_id = " & rP!Pat_id & " AND ((name Like '%Brief an %Dr%' Or name Like '%Arztbrief%' Or name Like 'Brief an HA%' Or name Like 'Brief an HAe%' Or name like 'Brief an %') AND name NOT LIKE '%Entwurf%')"
   Call rB.Open(sql, DBCn, adOpenDynamic, adLockReadOnly)
   If rB.BOF Then
    Call tubriefStandalone(rP!Pat_id, True)
   End If
  End If
weiter:
  rP.Move 1
 Loop
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doRestlicheBriefe/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' doRestlicheBriefe

Sub doRestlicheBriefe(frm As Lese, ab&)
 Dim sql$, sql1$
 Dim rP As New ADODB.Recordset
 Dim rB As New ADODB.Recordset
 Dim Zp As Date, Quartal$
 On Error GoTo fehler
 Zp = Now - 30
 Quartal = ZQuart(Zp)
' ab = InputBox("Ab welchem Pat. anfangen (0 = von vorne)?", "Rückfrage", 0)
 Call rP.Open("SELECT DISTINCT pat_id FROM `faelle` WHERE quartal = '" & Quartal & "' AND pat_id >= " & ab & " ORDER BY pat_id ", DBCn, adOpenDynamic, adLockReadOnly)
 Do While Not rP.EOF
  Set rB = Nothing
  sql = "SELECT * FROM `briefe` WHERE pat_id = " & rP!Pat_id & " AND zeitpunkt > " & DatFor_k(QAnf(Quartal)) & " AND ((name Like '%Brief an %Dr%' Or name Like '%Arztbrief%' Or name Like 'Brief an HA%' Or name Like 'Brief an HAe%') AND name NOT LIKE '%Entwurf%')"
  Call rB.Open(sql, DBCn, adOpenDynamic, adLockReadOnly)
  If rB.BOF Then
   Call tubriefStandalone(rP!Pat_id, True)
  End If
  rP.Move 1
 Loop
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doRestlicheBriefe/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' doRestlicheBriefe

Sub FaxSend(docName$, RecName$, RecNum$)
Dim FaxServer As New FAXCOMLib.FaxServer
Dim FaxDoc As New FAXCOMLib.FaxDoc
Dim FaxTiff As New FAXCOMLib.FaxTiff
Dim strFaxJob As FAXCOMLib.FaxJobs
Dim strFaxStatus As FAXCOMLib.FaxJob
Dim strFaxTiff As FAXCOMLib.FaxTiff
Dim strJobID&
On Error GoTo fehler
 Err.Clear
 FaxServer.connect ("anmeld2")

Set FaxDoc = FaxServer.CreateDocument(FaxSendDatei)
   
    FaxDoc.BillingCode = "Rechnungsnummer 381"
    FaxDoc.CoverpageName = vNS
    FaxDoc.CoverpageNote = "Insulinanaloga"
    FaxDoc.CoverpageSubject = "Insulinanaloga"
    FaxDoc.DiscountSend = 0
    FaxDoc.DisplayName = "G.Schade"
    FaxDoc.EmailAddress = "diabetologie@dachau-mail.de"
    FaxDoc.FaxNumber = RecNum
    FaxDoc.RecipientAddress = vNS
    FaxDoc.RecipientCity = vNS
    FaxDoc.RecipientCompany = "Praxis"
    FaxDoc.RecipientCountry = "D"
    FaxDoc.RecipientDepartment = vNS
    FaxDoc.RecipientHomePhone = vNS
    FaxDoc.RecipientName = RecName
    FaxDoc.RecipientOffice = vNS
    FaxDoc.RecipientOfficePhone = vNS
    FaxDoc.RecipientState = "Bayern"
    FaxDoc.RecipientTitle = vNS
    FaxDoc.RecipientZip = vNS
    FaxDoc.SendCoverpage = 0
    FaxDoc.SenderAddress = "Mittermayerstraße 13"
    FaxDoc.SenderCompany = "Diabetologische Schwerpunktpraxis"
    FaxDoc.SenderDepartment = "Schreibbüro"
    FaxDoc.SenderFax = "08131 616381"
    FaxDoc.SenderHomePhone = "616380"
    FaxDoc.SenderName = "Gerald Schade"
    FaxDoc.SenderOffice = "Praxis"
    FaxDoc.SenderOfficePhone = "616380"
    FaxDoc.SenderTitle = vNS
    FaxDoc.ServerCoverpage = 1
    FaxDoc.Filename = docName
    FaxDoc.DisplayName = docName
    strJobID = FaxDoc.send
'    MsgBox FaxServer.ArchiveDirectory
  
Set strFaxJob = FaxServer.GetJobs()
Set strFaxStatus = strFaxJob.Item(1)
    
On Error Resume Next

Set FaxServer = Nothing
Set FaxDoc = Nothing
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FaxSend/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' FaxSend

Public Function do_DMPAusgebStandAlone(Pat_id&, Optional fax1$, Optional Adressat$)
 Dim dc, VorString$, docName$, DT As DMPClass
 Dim mR1, mR2, mR3
 Dim rsAnam As New ADODB.Recordset
 Dim sverz$
 If fax1$ = vNS Then
  sverz = pVerz
 Else
  sverz = pVerz & "\zufaxen\2145\"
 End If
 On Error GoTo fehler
' dtbInit
' Set rsAnam = TabÖff("Anamnesebogen", "Pat_id")
' rsAnam.Open "SELECT * FROM `anamnesebogen` WHERE pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly ' 22.3.12: nach untens verschoben, da evtl. während des Aufrufs von DMPString Verlorengehen möglich
' rsAnam.Seek "=", Pat_id
 VorString = vNS
 nzw = vbCr
' If rsAnam!DMPhier <> CDate(0) AND rsAnam!HAimDMP = "Hausarzt im DMP" Then
'  VorString = "Bei mir fand am " + format$(rsAnam!DMPhier, "dd/mm/yyyy") + " die Einschreibung ins DMP statt. Diese sollte jedoch auf Sie als Hausarzt übertragen werden." + nzw
' End If
vorgetword:
 GetWord
 With Wapp
  .Options.SmartCutPaste = False
  On Error Resume Next
  .Options.SmartParaSELECTion = False
  On Error GoTo fehler
  Call .documents.Add(DMPVorlage)
  Set dc = .activedocument
  Dim VorZahl&, runde%
'  GoTo vorgetword
  FNr = 719
  VorZahl = dc.Range.End
  FNr = 1
  dc.Range.Insertafter VorString & Replace$(DMPString$(Pat_id, DT, , True), vbCrLf & vbCrLf, vbCrLf)
       Set mR1 = dc.content
       With mR1.Find
         .clearformatting
         .Text = " ("
         .replacement.Text = vNS
         .wrap = wdFindContinue
         .Format = False
         .Execute
       End With
       If mR1.Find.found Then
        Set mR2 = dc.Range(mR1.Start, mR1.Start)
        mR2.Find.Text = ")"
        mR2.Find.Execute
        If mR2.Find.found Then
         Set mR3 = dc.Range(mR1.Start, mR2.End)
         mR3.Font.Hidden = True
        End If
       End If
        
       Set mR3 = dc.Range(VorZahl, dc.Range.End)
      On Error Resume Next
       Dim Para
       Set Para = mR3.paragraphs.First.Range
       Do While Err.Number = 0
        With dc.Range(Para.Start, Para.Start + InStr(Para.Text, ":")).Font
         .Italic = True
         .Bold = True
        End With
        Set Para = Para.paragraphs.First.Next.Range
       Loop
  On Error GoTo fehler
  Dim SectZ%
  SectZ = dc.sections.COUNT
  dc.sections(SectZ).Range.ParagraphFormat.TabStops.ClearAll
  dc.sections(SectZ).Range.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(5.5), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderDots
  dc.sections(SectZ).Range.ParagraphFormat.FirstLineIndent = -20
  dc.sections(SectZ).Range.ParagraphFormat.LeftIndent = 20
  dc.Range(dc.Range.End - 1, dc.Range.End - 1).InsertBreak Type:=3 'wdSectionBreakContinuous
  dc.sections.Last.Range.PageSetup.TextColumns.Add Width:=CentimetersToPoints(11.4), Spacing:=CentimetersToPoints(0.6), EvenlySpaced:=False
'  dc.PageSetup.TopMargin = CentimetersToPoints(0.5)
'  dc.PageSetup.BottomMargin = CentimetersToPoints(0.5)

  Dim rs As New ADODB.Recordset
  Set rs = Nothing
'  sql = "SELECT `medarten`.*, `medplan`.medikament AS mmedikament,bemerkung,mo,mi,nm,ab,zn,--bbed AS j_bbed FROM `medplan` LEFT JOIN `medarten` ON `medplan`.medanfang = `medarten`.medikament WHERE `medplan`.mpnr = (SELECT MAX(mpnr) FROM `medplan` WHERE pat_id = " & Pat_id & " AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM `medplan` WHERE pat_id = " & CStr(Pat_id) & ")) AND not ISNULL(`medplan`.medikament) AND `medplan`.medikament <> '' AND `medplan`.pat_id = " & Pat_id
' `medplan`.medikament AS mmedikament,mo,mi,nm,ab,zn
  sql = "SELECT date(zeitpunkt) AS zp FROM `medplan` LEFT JOIN `medarten` ON `medplan`.medanfang = `medarten`.medikament WHERE `medplan`.pat_id = " & Pat_id & " AND `medplan`.mpnr = (SELECT MAX(mpnr) FROM `medplan` WHERE pat_id = " & Pat_id & " AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM `medplan` WHERE pat_id = " & CStr(Pat_id) & ")) AND not ISNULL(`medplan`.medikament) AND `medplan`.medikament <> '' AND `medplan`.pat_id = " & Pat_id & " GROUP BY zeitpunkt"
  rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  If Not rs.BOF Then
   sql = "SELECT left(`medplan`.medikament,16) AS `Med." & Format(rs!Zp, "d.m.yy") & ":`,mo,mi,nm,ab,zn FROM `medplan` LEFT JOIN `medarten` ON `medplan`.medanfang = `medarten`.medikament WHERE `medplan`.Pat_id = " & Pat_id & " AND `medplan`.mpnr = (SELECT MAX(mpnr) FROM `medplan` WHERE pat_id = " & Pat_id & " AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM `medplan` WHERE pat_id = " & CStr(Pat_id) & ")) AND not ISNULL(`medplan`.medikament) AND `medplan`.medikament <> '' AND `medplan`.pat_id = " & Pat_id
   Set rs = Nothing
   rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
   dc.Range.Insertafter vbCrLf & TabAusgeb(rs, Lese, , Chr$(11), , , , , , , True).Value
   Set mR2 = dc.paragraphs.Last.Range
   mR2.Find.Text = ":"
   mR2.Find.Execute
   If mR2.Find.found Then
    dc.Range(dc.paragraphs.Last.Range.Start, mR2.End).Bold = True
    dc.Range(dc.paragraphs.Last.Range.Start, mR2.End).Underline = 1
   End If
   With dc.paragraphs.Last.Range.Font
        .name = "Courier New"
        .size = 8
'        .Bold = False
        .Italic = False
'        .Underline = 0 ' wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .Strikethrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = False
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 0
        .Animation = 0 ' wdAnimationNone
    End With
    With dc.paragraphs.Last.Range.ParagraphFormat
        .LeftIndent = CentimetersToPoints(0)
        .RightIndent = CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = 0 ' wdLineSpaceSingle
        .Alignment = 0 ' wdAlignParagraphLeft
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = CentimetersToPoints(0)
        .OutlineLevel = 10 ' wdOutlineLevelBodyText
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
  End If
  dc.Range(dc.Range.End - 1, dc.Range.End - 1).InsertBreak Type:=8 'wdColumnBreak
  dc.Range(dc.Range.End - 1, dc.Range.End - 1).Insertafter "Diagnosen:" & vbCrLf
  dc.Range(dc.paragraphs.Last.Previous.Range.Start, dc.paragraphs.Last.Previous.Range.End - 1).Bold = True
  dc.Range(dc.paragraphs.Last.Previous.Range.Start, dc.paragraphs.Last.Previous.Range.End - 1).Underline = 1 ' wdUnderlineSingle
  Dim DiagTab() As CString
  dc.Range(dc.Range.End - 1, dc.Range.End - 1).Insertafter DiagString$(CStr(Pat_id), DiagTab, Now() - 180, True)
  rsAnam.Open "SELECT * FROM `anamnesebogen` WHERE pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
  docName = sverz + rsAnam!Nachname & " " & rsAnam!Vorname & ", PID " & rsAnam!Pat_id & ", DMP-Daten vom " + Format$(Now, "DD/MM/YY hh.mm.ss") & IIf(Adressat <> "" And Not IsNull(Adressat), " für " & Adressat, "") & IIf(LenB(fax1) = 0, vNS, " an Fax " & fax1) & ".doc"
  dc.SaveAs Filename:=Replace(Replace$(docName, "/", vNS), "\'", "'")
  dc.Close
'  .Visible = True
  
'  .Application.WindowState = wdWindowStateMaximize
'  .Activate
 End With
 do_DMPAusgebStandAlone = docName
 Exit Function
fehler:
 If FNr = 719 Or FNr = 1 Then
  If runde < 5 Then
   runde = runde + 1
   Resume vorgetword
  End If
 End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_DMPAusgebStandAlone/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_DMPAusgebStandalone

Public Function ZQSort$(Datum As Date) ' Für Abfragen mit Fallzuordnung
ZQSort = Year(Datum)
Select Case Month(Datum)
 Case 1, 2, 3: ZQSort = ZQSort & "1"
 Case 4, 5, 6: ZQSort = ZQSort & "2"
 Case 7, 8, 9: ZQSort = ZQSort & "3"
 Case Else:    ZQSort = ZQSort & "4"
End Select
End Function ' ZQuart$(Datum AS Date) ' Für Abfragen mit Fallzuordnung

Public Function ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung, 4 x so schnell wie untiges, mit CString hingegen 6 x so langsam wie dieses
Select Case Month(Datum)
 Case 1, 2, 3: ZQuart = "1"
 Case 4, 5, 6: ZQuart = "2"
 Case 7, 8, 9: ZQuart = "3"
 Case Else:    ZQuart = "4"
End Select
ZQuart = ZQuart & Year(Datum)
End Function ' ZQuart$(Datum AS Date) ' Für Abfragen mit Fallzuordnung

Public Function NQuart(Datum As Date) As Date ' nächstes Quartal
Dim Jahr%, Monat%
Jahr = Year(Datum)
Select Case Month(Datum)
 Case 1, 2, 3: Monat = 4
 Case 4, 5, 6: Monat = 7
 Case 7, 8, 9: Monat = 10
 Case Else:    Monat = 1: Jahr = Jahr + 1
End Select
NQuart = CDate("1." & Monat & "." & Jahr)
End Function ' ZQuart$(Datum AS Date) ' Für Abfragen mit Fallzuordnung

'Public Function ZQuart$(Datum AS Date) ' Für Abfragen mit Fallzuordnung
'Dim j AS String * 4, q AS String * 1
'On Error GoTo fehler
'j = Year(Datum)
'SELECT CASE Datum
' Case Is < CDate("1.4." + j): q = "1"
' Case Is < CDate("1.7." + j): q = "2"
' Case Is < CDate("1.10." + j): q = "3"
' Case Else: q = "4"
'End SELECT
'ZQuart = q + j
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 Then
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#End If
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZQuart/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End Function ' ZQuart$(Datum AS Date) ' Für Abfragen mit Fallzuordnung
Public Function zqtest()
 Dim i&, T1#, T2#, var, aktdat As Date
 aktdat = Now()
 T1 = Timer
 For i = 1 To 100000
  var = ZQuart(aktdat)
  aktdat = aktdat - 1
 Next i
 T2 = Timer
 Debug.Print T2 - T1
End Function ' zqtest

Function GesNamFn$(ByVal rs) ' AS DAO.Recordset) ' s.a. GesName(
 On Error GoTo fehler
   GesNamFn = rs!Nachname & ", " & nz(rs!Titel,vNS) + IIf(IsNull(rs!Titel) Or LenB(rs!Titel) = 0, vNS, ", ") + rs!Vorname
   On Error Resume Next
   GesNamFn = rs!NVorsatz + IIf(LenB(rs!NVorsatz) = 0, vNS, " ") + GesNamFn
   On Error GoTo fehler
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GesNam/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GesName

Function dtyp$(Inh$)
Select Case LCase(Inh)
'1','2','-','?','s' = sekundär,'g' = Gestationsdiabetes, 'p' = pathologische Glucosetoleranz oder gestörte Nüchternglucose
 Case "s":  dtyp = "sekundär"
 Case "g": dtyp = "Gestationsdiabetes"
 Case "p": dtyp = "'pathologische Glucosetoleranz oder gestörte Nüchternglucose'"
 Case Else: dtyp = Inh
End Select
End Function ' dtyp

Function DSeit(rsAnam As ADODB.Recordset, Optional objahr%)
 Dim rsDS As New ADODB.Recordset
 On Error GoTo fehler
    If IsNull(rsAnam("Diabetes seit")) Then
      DSeit = "?"
    ElseIf LCase(rsAnam("Diabetes seit")) = "bu" Then
'      DSeit = format$(rsAnam!Vorgestellt, "mm\/yy")
'      DSeit = format$(Dtb.OpenRecordset("SELECT fanf FROM `faelle` WHERE fid = (SELECT min(fid) FROM `faelle` WHERE pat_id = " + CStr(rsAnam!Pat_id) + ")")!Fanf, "mm\/yy")
      rsDS.Open "SELECT fanf FROM `faelle` WHERE fid = (SELECT min(fid) FROM `faelle` WHERE pat_id = " + CStr(rsAnam!Pat_id) + ")", DBCn, adOpenStatic, adLockReadOnly
      If Not rsDS.EOF Then
       DSeit = Format$(rsDS!Fanf, IIf(objahr, "yyyy", "mm\/yy"))
      End If
    Else: DSeit = CStr(rsAnam("Diabetes seit"))
    End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DSeit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DSeit(rsAnam)

Function TabPr(s1, s2)
'Dim nzl$
'nzl = vbcr + vblf ' vbcrlf kannte ich da noch nicht
#If False Then
 DmPStrS = DmPStrS & vbCrLf & s1
 If Len(s1) <= BreiteSp1 Then
  DmPStrS = DmPStrS & Left$(" ........................................", BreiteSp1 - Len(s1) - 1) & " "
 End If
 DmPStrS = DmPStrS & nz(s2,vNS) ' "........................"
#Else
 DmPStrS = DmPStrS & vbCrLf & s1 & Chr$(9) & nz(s2,vNS)
#End If
End Function ' TabPr(S1, S2)

'obpositiv
Function obPosi(s) As Boolean
 If IsNull(s) Then
  obPosi = False
 ElseIf VarType(s) = vbBoolean Then
  obPosi = s
 Else
  Dim SStr$
  SStr = s
  obPosi = Not obNein(SStr)
  If obPosi Then If SStr = "?" Then obPosi = False
 End If
End Function ' obPosi(S) AS Boolean
Public Function neuTher()
 Dim rNa As New ADODB.Recordset
 Call Lese.ProgStart
 rNa.Open "SELECT pat_id,gesname(pat_id) name FROM namen", DBCn, adOpenStatic, adLockReadOnly
 Do While Not rNa.EOF
  Debug.Print rNa!Pat_id, rNa!name
  DBCn.Execute "SELECT therartn(" & rNa!Pat_id & ")"
  rNa.MoveNext
  DoEvents
 Loop
End Function ' neuTher()

' wird in TherArt() aufgerufen, wo anhand des Diabetestyps und von Formularen Pumpentherapien ermittelt werden, sodann hiermit die übrigen Therapieformen
' parallel zu ThaFestleg
' aufgerufen in DMPString, Epikrise (2x) und Labordateianzeig
Function TherAuskunft(ByVal Pat_id$, ByVal obanf%, Optional ByRef insz%, Optional ByVal VorDat As Date, _
 Optional ByRef obIns As Boolean, Optional ByRef obAnal As Boolean, Optional ByRef obGlib As AntidiabMedType, _
 Optional ByRef obmetf As AntidiabMedType, Optional ByRef obGlucI As AntidiabMedType, _
 Optional ByRef obSHGlin As AntidiabMedType, Optional ByRef obGlit As AntidiabMedType, _
 Optional ByRef obdpp4 As AntidiabMedType, Optional ByRef obglp1 As AntidiabMedType, _
 Optional ByRef obsglt2 As AntidiabMedType, Optional ByRef obSonstAD As AntidiabMedType, _
 Optional ByRef obHMG As Boolean, Optional ByRef obAntihyp As Boolean, Optional ByRef obACEH As Boolean, _
 Optional ByRef obBetabl As Boolean, Optional ByRef obThro As Boolean, Optional ByRef obAntikoag%, _
 Optional ByVal Qmax$, Optional ByVal obRezIns As Boolean, Optional ByVal DT$, Optional ByRef obDiur%, _
 Optional ByRef obAT1%) As TherapieArt
' Dim rMA AS DAO.Recordset
 Dim obNIns% 'ob Normalinsulin
 Dim raMa As New ADODB.Recordset
 Dim AktMz%, medi$, rAF&
 Dim MPNr&(), MP0&, MPe&, mpz&, i&, MPNrl& ' erste und letzte, laufende MedPlan-Nummer des Quartals
 On Error GoTo fehler
 Call Lese.ProgStart
' QuartalMax: Falls befüllt, so wird die höchstwertige Therapie des in der Variablen angegebenen Quartals ermittelt
' Set rMa = TabÖff("MedArten", "Medikament")
'Medikamente:
 mpz = 0
 ReDim MPNr(mpz)
 If Qmax <> vNS Then
  MP0 = MedPlanNr(Pat_id, obAkt:=True, VorDat:=QAnf(Qmax), NurNr:=True)
  MPe = MedPlanNr(Pat_id, obAkt:=True, VorDat:=QEnd(Qmax) + 1, NurNr:=True)
  Set raMa = Nothing
  Call raMa.Open("SELECT DISTINCT mpnr FROM `medplan` WHERE pat_id = " & Pat_id & " AND mpnr >= " & MP0 & " AND mpnr <= " & MPe & " ORDER BY mpnr", DBCn, adOpenStatic, adLockReadOnly)
  Do While Not raMa.EOF
   mpz = mpz + 1
   ReDim Preserve MPNr(mpz)
   MPNr(mpz) = raMa!MPNr
   raMa.Move 1
  Loop
 Else
  mpz = 1
  ReDim MPNr(mpz)
  If obanf Then
   Call MedPlanAusAna(Pat_id)
  ElseIf obanf = 1 Then
   MPNr(0) = MedPlanNr(Pat_id, -1)
  Else
   MPNr(0) = MedPlanNr(Pat_id, Not obanf, VorDat)
  End If
 End If
 
 For i = 1 To mpz 'MPNrl = MPNrA To MPNrE
  If Qmax <> "" Then
   Call MedPlanNr(Pat_id, True, , , nr:=MPNr(i)) ' dann die übrigen Variablen hier jeweils neu befüllen
  End If
  obGlib = adnein: obmetf = adnein: obGlucI = adnein: obSHGlin = adnein: obGlit = adnein: obdpp4 = adnein: obglp1 = adnein: obsglt2 = adnein: obSonstAD = adnein
  obIns = 0: obAnal = 0: obHMG = 0: obAntihyp = 0: obACEH = 0: obBetabl = 0: obThro = 0
  insz = 0
  obNIns = 0
  For AktMz = 0 To MedZahl
   Dim DosH%
   DosH = DosHäuf(AktMz)
'  If DosH > 0 Then ' wenn auch bei Dosierung was drinsteht
   medi = Trim$(LCase$(Med(AktMz)))
   If LenB(medi) <> 0 And medi <> "-" Then
'    dim lzpos%
'    lzpos = InStr(medi, " ")
'    If lzpos > 0 Then medi = Left(medi, lzpos - 1)
    medi = GetMed(medi, 0)
'    rMA.Seek "=", medi
    Set raMa = Nothing
    raMa.Open "SELECT -hmg j_hmg, -hypt j_hypt, -aceh j_aceh, -betabl j_betabl, -thro j_thro, -antikoag j_antikoag, -glib j_glib, -metf j_metf, -gluci j_gluci, -shglin j_shglin, -glit j_glit, -dpp4 j_dpp4, -glp1 j_glp1, -sglt2 j_sglt2, -sonstad j_sonstad, -ins j_ins, -AnAl j_anal, -Diur j_diur, -AT1 j_at1, m.* FROM `medarten` m WHERE medikament = '" & Replace$(medi, "'", "''") & "'", DBCn, adOpenDynamic, adLockReadOnly
    If raMa.BOF Then
     If DosH > 0 Then
'      MsgBox "Medikament: " + Medi + " noch nicht in Tabelle MedArten erfaßt"
      Debug.Print "Medikament: " + medi + " noch nicht in Tabelle `medarten` erfaßt"
      InsKorr DBCn, DBCnS, "INSERT INTO `medarten`(langname,medikament,hinzugefügt,pat_id) VALUES('" & Med(AktMz) & "','" & UCase$(medi) & "'," & DatFor_k(Now) & "," & Pat_id & ")", rAF
     End If
    Else
     If Not IsNull(raMa!j_antikoag) Then If raMa!j_antikoag <> 0 Then obAntikoag = True ' Bei Marcumar steht meist keine Dosierung drin
     If DosH > 0 Then
      If Not IsNull(raMa!j_glib) Then If raMa!j_glib <> 0 Then obGlib = adja
      If Not IsNull(raMa!j_Metf) Then If raMa!j_Metf <> 0 Then obmetf = adja
      If Not IsNull(raMa!j_GlucI) Then If raMa!j_GlucI <> 0 Then obGlucI = adja
      If Not IsNull(raMa!j_SHGlin) Then If raMa!j_SHGlin <> 0 Then obSHGlin = adja
      If Not IsNull(raMa!j_Glit) Then If raMa!j_Glit <> 0 Then obGlit = adja
      If Not IsNull(raMa!j_dpp4) Then If raMa!j_dpp4 <> 0 Then obdpp4 = adja
      If Not IsNull(raMa!j_glp1) Then If raMa!j_glp1 <> 0 Then obglp1 = adja
      If Not IsNull(raMa!j_sglt2) Then If raMa!j_sglt2 <> 0 Then obsglt2 = adja
      If Not IsNull(raMa!j_SonstAD) Then If raMa!j_SonstAD <> 0 Then obSonstAD = adja
      If Not IsNull(raMa!j_InS) And (raMa!j_InS <> 0 Or (Not raMa!j_AnAl <> 0 And InStrB(LCase(medi), "insulin") > 0)) Then
        obIns = True
        insz = insz + DosH
      End If
      If Not IsNull(raMa!j_AnAl) And raMa!j_AnAl <> 0 Then
        obAnal = True
        insz = insz + DosH
      End If
      If raMa!insart = 1 Then
       obNIns = True
      End If
      If Not IsNull(raMa!j_hmg) Then If raMa!j_hmg <> 0 Then obHMG = True
      If Not IsNull(raMa!j_hypt) Then If raMa!j_hypt <> 0 Then obAntihyp = True
      If Not IsNull(raMa!j_diur) Then If raMa!j_diur <> 0 Then obDiur = True
      If Not IsNull(raMa!j_betabl) Then If raMa!j_betabl <> 0 Then obBetabl = True
      If Not IsNull(raMa!j_aceh) Then If raMa!j_aceh <> 0 Then obACEH = True
      If Not IsNull(raMa!j_at1) Then If raMa!j_at1 <> 0 Then obAT1 = True
      If Not IsNull(raMa!j_thro) Then If raMa!j_thro <> 0 Then obThro = True
     End If ' dosh > 0
    End If ' rma.nomatch
   End If
'  End If 'doshäuf(aktmz)
  Next AktMz ' aktmz = 0 to mednr
  
#If thaalt Then
  TherAuskunft = offen
  Dim Thfakt As TherapieArt
  Thfakt = offen
  If obNIns Then
    Thfakt = ict
  Else
    Select Case insz
     Case 0
      If obRezIns Then
       Thfakt = ict
      Else
       If obglp1 = adja Then
        Thfakt = glp1
       ElseIf obGlib = adja Or obmetf = adja Or obGlucI = adja Or obSHGlin = adja Or obGlit = adja Or obdpp4 = adja Or obsglt2 = adja Or obSonstAD = adja Then
        Thfakt = OAD
       Else
        If DT = "1" Then
         Thfakt = ct
        Else
         Thfakt = Diät
        End If
       End If
      End If
     Case 1
      If obglp1 = adja Then
       Thfakt = glp1ins
      ElseIf obGlib = adja Or obmetf = adja Or obGlucI = adja Or obSHGlin = adja Or obGlit = adja Or obdpp4 = adja Or obsglt2 = adja Or obSonstAD = adja Then
       Thfakt = komb
      Else
       Thfakt = ct
      End If
     Case 2
      Thfakt = ct
     Case Is >= 3
      Thfakt = ict
    End Select
   End If
   If Thfakt > TherAuskunft Then
    TherAuskunft = Thfakt
   End If
#End If

 Next i
' raMa.Close
Exit Function
fehler:
Dim tonRunde%
'For tonRunde = 1 To 10
' Call Sound(WinDir + "\media\Windows XP-Standard.wav")
'Next tonRunde
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in TherAuskunft/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' TherAuskunft(pat_id, obanf%, insz, Optional obIns, Optional obAnal, Optional obGlib, Optional obMetf, Optional obGlucI, Optional obSHGlin, Optional obGlit, Optional obSonstAD, Optional obHMG, Optional obAntihyp, Optional obThro)

Function diI%(icd_str$, Optional pid, Optional abDat As Date, Optional auchZ As Boolean, Optional Weiteres$, Optional ByRef ICD$, Optional ByRef diagdat As Date)
 Dim raDT As New ADODB.Recordset, raFa As New ADODB.Recordset
 Dim diIsql$
 On Error GoTo fehler
 If Not IsMissing(pid) Then Pat_id = CLng(pid)
 If Pat_id = 0 Then
  MsgBox "Stop in diI durch Pat_id = 0"
  Stop
  Exit Function
 End If
 Dim icdsql$
 If InStrB(icd_str, " ") = 0 Then
  icdsql = "icd like '" & icd_str & "%'"
 Else
 ' ergänzt 13.8.15
  Dim SpliZ&, i&, splerg$()
  SpliZ = SplitNeu(icd_str, " ", splerg)
  icdsql = "("
  For i = 0 To SpliZ - 1
   icdsql = icdsql & "icd like '" & splerg(i) & "%'"
   If i = SpliZ - 1 Then icdsql = icdsql & ")" Else icdsql = icdsql & " OR "
  Next i
 End If
 diIsql = "SELECT " & IIf(lies.obMySQL, vNS, "top 1 ") & "* FROM `diagnosen` WHERE pat_id = " & Pat_id & " AND " & icdsql & " AND diagsicherheit in ('G','V',''" & IIf(auchZ, ",'Z'", "") & ")  AND COALESCE(f6010,0)=0 " & IIf(abDat <> 0, " AND diagdatum >= " & DatFor_k(abDat), "") & IIf(LenB(Weiteres) <> 0, " AND " & Weiteres, "")
 Dim lddat As Date
 raFa.Open "SELECT MAX(bhfb) AS lddat FROM `faelle` WHERE pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
 If Not raFa.BOF Then
  lddat = nz(raFa!lddat,0)
  diIsql = diIsql & " AND (obdauer <> 0 Or diagdatum >= " & DatFor_k(lddat) & ")"
 End If
 If lies.obMySQL Then diIsql = diIsql & " LIMIT 1"
 Call raDT.Open(diIsql, DBCn, adOpenStatic, adLockReadOnly)
 diI = Not raDT.BOF
 If diI Then
  ICD = raDT!ICD
  diagdat = raDT!DiagDatum
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diI/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function 'diI(icd_str, PID) AS Boolean

' 9.9.15: Parameter nach PID eingebaut, obwohl diese bisher nicht gebraucht wurden
Function diT%(diagtxt$, Optional pid, Optional abDat As Date, Optional auchZ As Boolean, Optional Weiteres$, Optional ByRef ICD$, Optional ByRef diagdat As Date)
 Dim raDT As New ADODB.Recordset
 Dim raFa As New ADODB.Recordset
 Dim diTsql$
 On Error GoTo fehler
 If Not IsMissing(pid) Then Pat_id = pid
 If Pat_id = 0 Then
  MsgBox "Stop in diT durch Pat_id = 0"
  Stop
  Exit Function
 End If
 Dim dtxtsql$
 If InStrB(diagtxt, " ") = 0 Then
  dtxtsql = "diagtext like '%" & diagtxt & "%'"
 Else
 ' ergänzt 13.8.15
  Dim SpliZ&, i&, splerg$()
  SpliZ = SplitNeu(diagtxt, " ", splerg)
  dtxtsql = "("
  For i = 0 To SpliZ - 1
   dtxtsql = dtxtsql & "diagtext like '%" & splerg(i) & "%'"
   If i = SpliZ - 1 Then dtxtsql = dtxtsql & ")" Else dtxtsql = dtxtsql & " OR "
  Next i
 End If
 diTsql = "SELECT " & IIf(lies.obMySQL, vNS, "top 1 ") & "* FROM `diagnosen` WHERE pat_id = " & Pat_id & " AND " & dtxtsql & " AND diagsicherheit in ('G','V',''" & IIf(auchZ, ",'Z'", "") & ") AND COALESCE(f6010,0)=0 " & IIf(abDat <> 0, " AND diagdatum >= " & DatFor_k(abDat), "") & IIf(LenB(Weiteres) <> 0, " AND " & Weiteres, "")
 Dim lddat As Date
 raFa.Open "SELECT MAX(bhfb) AS lddat FROM `faelle` WHERE pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
 If Not raFa.BOF Then
  If Not IsNull(raFa!lddat) Then
   lddat = raFa!lddat
   diTsql = diTsql & " AND (obdauer <> 0 Or diagdatum >= " & DatFor_k(lddat) & ")"
  End If
 End If
 If lies.obMySQL Then diTsql = diTsql & " LIMIT 1"
 Call raDT.Open(diTsql, DBCn, adOpenStatic, adLockReadOnly)
 diT = Not raDT.BOF
 If diT Then
  ICD = raDT!ICD
  diagdat = raDT!DiagDatum
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diT/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function 'diT(icd_str, PID) AS Boolean

'Function Seite$()
'  SELECT CASE raDT!DiagSeite
'  Case "R": Seite = "rechts"
'  Case "L": Seite = "links"
'  Case "B": Seite = "bds."
' End SELECT
'End Function ' seite$()

Function PulsParse$(Feld, Optional repath As Boolean, Optional lipath As Boolean)
     Dim Add$, rE$, li$
     repath = False
     lipath = False
     Feld = tfeld(Feld)
     If Feld = "-" Then
      Add = "-/-"
     ElseIf Feld = "+" Or Feld = "n" Then
      Add = "+/+"
     ElseIf Feld = "(-)" Then
      Add = "(-)/(-)"
     ElseIf Feld = "(+)" Then
      Add = "(+)/(+)"
     ElseIf Feld = "++" Then
      Add = "++/++"
     ElseIf Feld = "--" Then
      Add = "--/--"
     ElseIf Feld = "?" Then
      Add = "?/?"
     ElseIf Feld = "(?)" Then
      Add = "(?)/(?)"
     Else: Add = CStr(Feld)
     End If
     If InStrB(Add, "/") = 0 Then Add = Replace$(Add, ",", "/")
     Add = Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Add, "biph.", "biphasisch "), "biphas.", "biphasisch "), "hyperä.", "hyperämisch "), "hyperä", "hyperämisch "), "mono", "monophasisch"), "pst", "poststenotisch ")
     If Right$(Add, 2) = "bi" Then Add = Add + "phasisch"
     If InStrB(Add, ",") = 0 And InStrB(Add, "/") = 0 And InStrB(Add, "bds") = 0 And InStrB(Add, "beid") = 0 Then
      Add = "bds. " + Add
     End If
     PulsParse = Add
     rE = Add
     li = Add
     If InStrB(Add, "/") <> 0 Then
      rE = Left(Add, InStr(Add, "/") - 1)
      li = Mid$(Add, InStr(Add, "/") + 1)
     End If
     If InStrB(Add, "|") <> 0 Then
      rE = Left(Add, InStr(Add, "|") - 1)
      li = Mid$(Add, InStr(Add, "|") + 1)
     End If
     If InStrB(rE, "-") <> 0 Or InStrB(rE, "?") <> 0 Then repath = True
     If InStrB(li, "-") <> 0 Or InStrB(li, "?") <> 0 Then lipath = True
End Function ' PulsParse

Function tfeld$(s)
If IsNull(s) Then
 tfeld = vNS
Else
 tfeld = s
End If
End Function ' TFeld$(S)

Function obMonPath(Feld, rep%, lip%)
Dim TestFeld$, Tre$, Tli$
rep = 1
lip = 1
SensSplit Feld, Tre, Tli
If InStrB(Tre, "/5") <> 0 Then
 Tre = Replace$(Tre, "/5", "")
 Tli = Replace$(Tli, "/5", "")
 If Tre = "" Or Tre = "~" Or InStrB(Tre, "5") <> 0 Then rep = 0
 If Tli = "" Or Tli = "~" Or InStrB(Tli, "5") <> 0 Then lip = 0
ElseIf InStrB(Tre, "/3") <> 0 Then
 Tre = Replace$(Tre, "/3", "")
 Tli = Replace$(Tli, "/3", "")
 If Tre = "" Or Tre = "~" Or InStrB(Tre, "3") <> 0 Then rep = 0
 If Tli = "" Or Tli = "~" Or InStrB(Tli, "3") <> 0 Then lip = 0
Else
 If Tre = "" Then rep = 0
 If Tli = "" Then lip = 0
End If
End Function ' obMonPath(Feld, rep%, lip%)

Function obKWPath(Feld, rep%, lip%)
Dim TestFeld$, Tre$, Tli$
rep = 1
lip = 1
SensSplit Feld, Tre, Tli
Tre = Replace$(Tre, "/5", "")
Tli = Replace$(Tli, "/5", "")
If Tre = "" Or Tre = "~" Or InStrB(Tre, "4") <> 0 Or InStrB(Tre, "5") <> 0 Then rep = 0
If Tli = "" Or Tli = "~" Or InStrB(Tli, "4") <> 0 Or InStrB(Tli, "5") <> 0 Then lip = 0
End Function ' obKWPath(Feld, rep%, lip%)
Function obVibPath(Feld, rep%, lip%, Optional obStreng%)
Dim TestFeld$, Tre$, Tli$
rep = 1
lip = 1
SensSplit Feld, Tre, Tli
Tre = Replace$(Tre, "/8", "")
Tli = Replace$(Tli, "/8", "")
If (Tre = "" Or Tre = "~" Or (obStreng = 0 And InStrB(Tre, "4") <> 0) Or InStrB(Tre, "5") <> 0 Or InStrB(Tre, "6") <> 0 Or InStrB(Tre, "7") <> 0 Or InStrB(Tre, "8") <> 0) And InStrB(Tre, ">8") = 0 Then rep = 0
If (Tli = "" Or Tli = "~" Or (obStreng = 0 And InStrB(Tli, "4") <> 0) Or InStrB(Tli, "5") <> 0 Or InStrB(Tli, "6") <> 0 Or InStrB(Tli, "7") <> 0 Or InStrB(Tli, "8") <> 0) And InStrB(Tli, ">8") = 0 Then lip = 0
End Function ' obVibPath(Feld, rep%, lip%)

Function SensSplit(SensStrg, rE$, li$)
If IsNull(SensStrg) Then SensStrg = ""
If InStrB(SensStrg, "|") <> 0 Then
      rE = LTrim$(Left$(CStr(SensStrg), InStr(CStr(SensStrg), "|") - 1))
      li = LTrim$(Mid$(CStr(SensStrg), InStr(CStr(SensStrg), "|") + 1))
ElseIf InStrB(SensStrg, ",") <> 0 Then
      rE = Left$(CStr(SensStrg), InStr(CStr(SensStrg), ",") - 1)
      li = Mid$(CStr(SensStrg), InStr(CStr(SensStrg), ",") + 1)
Else
      rE = SensStrg
      li = SensStrg
End If
rE = Trim$(rE)
li = Trim$(li)
End Function ' SensSplit(SensStrg, re$, li$)

' in Epikrise, rrpruef und DMPString
 Public Function GetPrRR$(rsAnahier As ADODB.Recordset, RRsyst%, RRdiast%, Optional obdiastkorr% = 0)
 '  Dim rrr AS DAO.Recordset
'  Set rrr = TabÖff("RR", "Auswahl")
'  rrr.Seek "=", CStr(rs!Pat_id)
  On Error GoTo fehler
  GetPrRR = vNS
'  rarr.Open "SELECT * FROM rr WHERE pat_id = " & rsAnahier!Pat_id & " ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
  Dim rarr As New ADODB.Recordset
  rarr.Open "SELECT rs,rd FROM (SELECT rrsyst(rr) rs, rrdiast(rr) rd, zeitpunkt FROM rr WHERE Pat_id = " & rsAnahier!Pat_id & ") i WHERE rs<>0 AND rd<>0 ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
  If rarr.EOF Then
   Set rarr = Nothing
   rarr.Open "SELECT rs,rd FROM (SELECT rrsyst(rr) rs, rrdiast(rr) rd, zeitpunkt FROM rr WHERE Pat_id = " & rsAnahier!Pat_id & ") i WHERE rs<>0 ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
  End If
  If Not rarr.BOF Then
   RRsyst = rarr!rs
   RRdiast = rarr!rD
   GetPrRR = RRsyst & "/" & RRdiast & " mm Hg"
  Else
   With rsAnahier
'    If Not rarr.EOF Then
'     GetPrRR = rarr!RR & IIf(InStrB(rarr!RR, "mm Hg") = 0 AND InStrB(rarr!RR, "mmHg") = 0, " mm Hg", "") + " (" + Format$(rarr!Zeitpunkt, "DD.MM.YY") + ")"
'    ElseIf
    If tfeld(!RR) <> "" Then
     GetPrRR = tfeld(!RR)
    ElseIf tfeld(!RRTurboMed) <> "" Then
     GetPrRR = tfeld(!RRTurboMed)
    ElseIf tfeld(!Blutdruckwerte) <> "" Then
     GetPrRR = tfeld(!Blutdruckwerte)
    End If
    If GetPrRR <> "" Then
     Call dodoRRParse(GetPrRR, RRsyst, RRdiast)
    End If
   End With
  End If
  If obdiastkorr Then If RRdiast < 30 Then RRdiast = Min(70, RRsyst) ' 31.10.20
  Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetPrRR/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GetPrRR

#If False Then
Function SchulzBest%(Pat_id&, zp1$, zpl$, Optional abDat)
  Dim sqls$, lapp As New ADODB.Recordset
  On Error GoTo fehler
  sqls$ = "SELECT * FROM `eintraege` WHERE pat_id = " & CStr(Pat_id) + " AND art = ""schul"" "
  If Not IsMissing(abDat) Then If Now > CDate("15.10.05") Then sqls = sqls + "and Zeitpunkt >= " & DatFor_k(abDat) & " "
  sqls = sqls + "ORDER BY zeitpunkt"
  lapp.Open sqls, DBCn, adOpenDynamic, adLockReadOnly
  If Not lapp.BOF Then
   lapp.MoveLast
   zpl = Format$(lapp!Zeitpunkt, "dd/mm/yy")
   lapp.MoveFirst
   zp1 = Format$(lapp!Zeitpunkt, "dd/mm/yy")
  End If
  lapp.Close
  lapp.Open "SELECT COUNT(0) AS zl FROM (" & sqls & ")" & IIf(lies.obMySQL, " AS innen", ""), DBCn, adOpenStatic, adLockReadOnly
  SchulzBest = lapp!zl
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SchutzBest/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'SchulzBest
#End If

Function rrEmpf$(RRsyst%, RRdiast%, Pat_id$, Optional freq%)
  Dim gwSyst%, gwDiast%, Verzei$, Datei$, taskid&
  Dim erg$
  Call Lese.ProgStart
  If Not diI("I10", Pat_id) Then
   If RRsyst > 149 Or RRdiast > 100 Then
'    If MsgBox("Achtung: Fehlt Diagnose Hypertonie bei Pat_id = " & CStr(Pat_id) & "?", vbYesNo) = vbYes Then
'     MsgBox "Bitte in Turbomed eintragen"
'    End If
    Verzei = pVerz & "plz\tmp\"
    Datei = Verzei & "Hypertoniehinweis Pat." & Pat_id & ".txt"
    VerzPrüf Verzei
    erg = Dir(Datei)
    If LenB(erg) = 0 Then
     Open Datei For Output As #371
     Print #371, "Pat." & Pat_id & " hat RR eingetragen von: " & RRsyst & "/" & RRdiast & "."
     Print #371, "   Diagnose I10 aber nicht gefunden."
     Print #371, "   " & Now()
     Close #371
    End If
    If freq = 0 Then freq = 600
    Call Beep(freq, 500)
    taskid = zeigan(Datei, vbNormalFocus)
    On Error Resume Next
    AppActivate taskid
    On Error GoTo 0
   End If
   rrEmpf = "halten"
  Else
   If RRsyst > 0 Then
    If diI("N08.3", Pat_id) And Not obkNeph Then
     gwSyst = 121
     gwDiast = 81
    Else
     gwSyst = 131
     gwDiast = 86
    End If
    If (RRsyst < gwSyst And RRdiast < gwDiast) Then
     rrEmpf = "halten"
    Else
     rrEmpf = "senken"
    End If
  End If
 End If
 End Function ' RREmpf

Function DatInStr(teststr$, Optional Jahr) As Date
 Dim buch As String * 1, SL%, aktb%, p1%, pL%, puZ%, ZwiStr$
 SL = Len(teststr)
 p1 = 0
 pL = 0
 puZ = 0
 For aktb = 1 To SL
  buch = Mid$(teststr, aktb, 1)
  If p1 = 0 Then
   If InStrB("0123456789", buch) <> 0 And buch <> "" Then p1 = aktb
  End If
  If p1 > 0 Then
   If buch = "." Then puZ = puZ + 1
   If InStrB("0123456789" + IIf(puZ > 2, vNS, "."), buch) = 0 Then
    pL = aktb - p1
    Exit For
   End If
  End If
 Next
 DatInStr = CDate(0)
 If p1 > 0 And pL > 0 Then
  ZwiStr = Mid$(teststr, p1, pL)
  If Not IsDate(ZwiStr) Then
   ZwiStr = ZwiStr + CStr(Jahr)
  End If
  If IsDate(ZwiStr) Then
   DatInStr = CDate(ZwiStr)
  End If
 End If
End Function ' DatInStr(TestStr$, Optional jahr) AS Date

Function MedPlanAusAna(ByVal Pat_id$)
 Dim rsNa As New ADODB.Recordset, i%, j%, inkl%, k%, dosstr$, dosier$()
 
' Set rsna = TabÖff("Anamnesebogen", "Pat_id")
 rsNa.Open "SELECT * FROM `anamnesebogen` WHERE pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
' rsNa.Seek "=", Pat_id
 MedZahl = 0
 If rsNa.EOF Then Exit Function
 MedZahl = 3
 ReDim Med$(MedZahl)
 ReDim Dos$(5, MedZahl)
 j = 0
 For i = 1 To 4
  With rsNa.Fields("DiabetesMedikament " + CStr(i))
   If IsNull(.Value) Then
    MedZahl = MedZahl - 1
    GoTo weiter
   End If
   Med(j) = .Value
   With rsNa.Fields("DiabetesMedikament " + CStr(i) + " Menge")
    If Not IsNull(.Value) Then
     If Len(.Value) > 0 Then
      dosstr = vNS
      inkl = 0
      For k = 1 To Len(.Value)
       If Mid$(.Value, k, 1) = "-" And inkl Then
        dosstr = dosstr + "€"
       Else
        dosstr = dosstr + Mid$(.Value, k, 1)
       End If
       If Mid$(.Value, k, 1) = "(" Then inkl = -1
       If Mid$(.Value, k, 1) = ")" Then inkl = 0
      Next
     End If
     dosier = Split(dosstr, "-")
     For k = 0 To UBound(dosier)
      dosier(k) = Replace$(dosier(k), "€", "-")
     Next k
     If UBound(dosier) > -1 Then Dos(0, j) = dosier(0)
     If UBound(dosier) > 0 Then Dos(1, j) = dosier(1)
     Select Case UBound(dosier)
      Case 2
       Dos(3, j) = dosier(2)
      Case 3
       Dos(3, j) = dosier(2)
       Dos(4, j) = dosier(3)
      Case Is > 3
       Dos(2, j) = dosier(2)
       Dos(3, j) = dosier(3)
       Dos(4, j) = dosier(4)
     End Select
    End If
   End With
   j = j + 1
  End With
weiter:
 Next i
 If MedZahl = -1 Then
  ReDim Med$(0)
  ReDim Dos$(5, 0)
 Else
  ReDim Preserve Med$(MedZahl)
  ReDim Preserve Dos$(5, MedZahl)
 End If
End Function 'MedPlanAusAna

' Wählt den ersten oder den aktuellen Medikamentenplan aus obakt = true = aktuellen, obakt = false = ersten
Function MedPlanNr&(ByVal Pat_id$, ByVal obAkt, Optional ByVal VorDat As Date, Optional ByVal NurNr%, Optional ByVal nr&)
 Dim raFIM As New ADODB.Recordset
 On Error GoTo fehler
' sql1 = "SELECT " & IIf(obAkt, "max", "min") & "(mpnr) FROM `medplan` WHERE pat_id = " + CStr(Pat_id)
' If vordat <> CDate(0) Then sql1 = sql1 + " AND zeitpunkt < " & DatFor_k(vordat + 1)
' sql1 = sql1 + " ORDER BY zeitpunkt" + IIf(obAkt, " DESC", "") + ", Medikament, feldnr"
 'Set rFIM = Dtb.OpenRecordset(sql1, dbOpenDynaset)
 If nr <> 0 Then
  MedPlanNr = nr
 Else
'  sql1 = "SELECT " & IIf(obAkt, "max", "min") & "(mpnr) AS mpmax FROM `medplan` WHERE pat_id = " + CStr(Pat_id)
  sql1 = "SELECT " & IIf(obAkt, "max", "min") & "(mpnr) mpmax FROM medplan mp1 WHERE pat_id = " & Pat_id & " AND zeitpunkt = (SELECT " & IIf(obAkt, "max", "min") & "(zeitpunkt) FROM medplan mp0 WHERE pat_id = mp1.pat_id)"
  If VorDat <> CDate(0) Then sql1 = sql1 + " AND zeitpunkt < " & DatFor_k(VorDat + 1)
  raFIM.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
' If Not ISNULL(raFIM.Fields(0)) Then
  If Not raFIM.BOF And Not IsNull(raFIM!mpmax) Then
   MedPlanNr = raFIM!mpmax
   If NurNr Then Exit Function
  Else
   Exit Function
  End If
 End If
 sql1 = "SELECT * FROM `medplan` WHERE pat_id = " & Pat_id & " AND mpnr = " & MedPlanNr
 Set raFIM = Nothing
 raFIM.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
 ' End If
  lBehDat = CDate(0)
  MedZahl = 0
  If Not raFIM.BOF Then
   raFIM.MoveFirst
   On Error GoTo fehler
   Do While Not raFIM.EOF
    If lBehDat <> CDate(0) And raFIM!Zeitpunkt <> lBehDat Then Exit Do
    If lBehDat = CDate(0) Then lBehDat = raFIM!Zeitpunkt
    If CLng(raFIM!FeldNr) > MedZahl Then MedZahl = CLng(raFIM!FeldNr)
    raFIM.Move 1
   Loop
   ReDim Med$(MedZahl)
   ReDim Dos$(5, MedZahl)
   MedZahl = 0
   raFIM.MoveFirst
   Do While Not raFIM.EOF
    If raFIM!Zeitpunkt <> lBehDat Then Exit Do
    On Error GoTo fehler
    MedNr = CLng(raFIM!FeldNr)
    If MedNr > MedZahl Then
     MedZahl = MedNr
     ReDim Preserve Med$(MedZahl)
     ReDim Preserve Dos$(5, MedZahl)
    End If
    If Not IsNull(raFIM!Medikament) Then
     Med(MedNr) = raFIM!Medikament
     Dos(0, MedNr) = nz(raFIM!mo,vNS)
     Dos(1, MedNr) = nz(raFIM!mi,vNS)
     Dos(2, MedNr) = nz(raFIM!nm,vNS)
     Dos(3, MedNr) = nz(raFIM!ab,vNS)
     Dos(4, MedNr) = nz(raFIM!Zn,vNS)
    End If
    raFIM.Move 1
   Loop
  End If ' not bof
' End If ' not ISNULL
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in MedPlanNr/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' MedPlanNr(Pat_id, obAkt)

Function DosHäuf%(MedNr%)
 Dim i%
 DosHäuf = 0
 If MedNr > -1 Then
  If MedNr > MedZahl Or MedZahl = 0 Then
   MedZahl = MedNr
   ReDim Preserve Med$(MedZahl)
   ReDim Preserve Dos$(5, MedZahl)
  End If
  For i = 0 To 4
   If Not IsEmpty(Dos(i, MedNr)) Then
    If Dos(i, MedNr) <> "" Then
     DosHäuf = DosHäuf + 1
     If IsNumeric(Dos(i, MedNr)) Then
      If CLng(Dos(i, MedNr)) = 0 Then DosHäuf = DosHäuf - 1
     End If
    End If
 '   If DosHäuf Then Exit Function
   End If
  Next
 End If ' MedNr> -1
End Function 'DosHäuf%(MedNr%)

#If doppelt Then
' 20.3.17: folgende Funktion sollte jetzt durch SQL-Funktionen RRsyst und RRdiast erübrigt sein
Public Function dodoRRParse(ByVal erg$, RRsyst%, RRdiast%, Optional Zp As Date)
  Dim i%, T1$, T1a$, pos%, runde%
  On Error GoTo fehler
  dodoRRParse = -1
  On Error GoTo fehler
  runde = 0
  Do
   runde = runde + 1
   RRsyst = 0
   RRdiast = 0
   Zp = 0
   pos = InStr(erg, "/")
   If pos > 1 Then
    T1 = Left(erg, pos - 1)
    If Len(T1) > 2 Then
     If Mid$(T1, Len(T1) - 2, 1) = "," And InStrB("0123456789", Mid$(T1, Len(T1) - 1, 1)) <> 0 Then T1 = Left(T1, Len(T1) - 2)
    End If
    T1a = Right$(T1, 3)
    If IsNumeric(T1a) Then
     RRsyst = Val(T1a)
    Else
     T1a = Right$(T1a, 2)
     If IsNumeric(T1a) Then
      RRsyst = Val(T1a)
     End If
    End If
    erg = Mid$(erg, pos + 1)
    T1a = Left(erg, 3)
    If IsNumeric(T1a) Then
     RRdiast = Val(T1a)
     erg = Mid$(erg, 4)
    Else
     T1a = Left(T1a, 2)
     If IsNumeric(T1a) Then
      RRdiast = Val(T1a)
      erg = Mid$(erg, 2)
     End If
    End If
    If InStr(erg, "(") > InStr(erg, "/") Then
     T1a = Mid$(erg, InStr(erg, "("))
     If InStr(T1a, ")") > 1 Then
      T1a = Left(T1a, InStr(T1a, ")") - 1)
      If IsDate(T1a) Then Zp = CDate(T1a)
     End If
    End If
   Else ' pos > 1
    pos = InStr(erg, "syst")
    If pos > 1 Then
     T1 = Left(erg, pos - 1)
     For i = Len(T1) To 0 Step -1
      If InStrB("0123456789", Mid$(T1, i, 1)) <> 0 And Mid$(T1, i, 1) <> "" Then Exit For
      T1 = Left(T1, Len(T1) - 1)
      If T1 = "" Then Exit For
     Next
     If Len(T1) > 2 Then
      If Mid$(T1, Len(T1) - 2, 1) = "," And InStrB("0123456789", Mid$(T1, Len(T1) - 1, 1)) <> 0 And Mid$(T1, Len(T1) - 1, 1) <> "" Then T1 = Left(T1, Len(T1) - 2)
     End If
     If T1 = vNS Then GoTo nix
     T1a = Right$(T1, 3)
     If IsNumeric(T1a) Then
      RRsyst = Val(T1a)
     Else
      T1a = Right$(T1a, 2)
      If IsNumeric(T1a) Then
       RRsyst = Val(T1a)
      End If
      If RRsyst < 50 Then
       T1a = Left(T1a, Len(T1a) - 2)
       For i = Len(T1a) To 0 Step -1
        If T1a = vNS Then Exit For
        If InStrB("0123456789", Mid$(T1a, i, 1)) <> 0 And Mid$(T1a, i, 1) <> vNS Then Exit For
        T1 = Left(T1a, Len(T1a) - 1)
       Next
       If Len(T1a) > 2 Then
        If Mid$(T1a, Len(T1a) - 2, 1) = "," And InStrB("0123456789", Mid$(T1a, Len(T1a) - 1, 1)) <> 0 And Mid$(T1a, Len(T1a) - 1, 1) <> "" Then T1a = Left(T1a, Len(T1a) - 2)
       End If
       T1a = Right$(T1a, 3)
       If IsNumeric(T1a) Then
        RRsyst = Val(T1a)
       End If
      End If
     End If
     erg = Mid$(erg, pos + 4)
    Else ' pos > 1
     erg = vNS
    End If ' pos > 1
   End If ' pos > 1
   If RRsyst > 0 Then Exit Do
   If erg = "" Then
nix:
     dodoRRParse = 0
     Exit Do
   End If
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dodoRRParse/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dodoRRParse
#End If

Function WieTabak(Pat_id&) As ZigSt
  Dim raZig As New ADODB.Recordset
  WieTabak = doTabakSt(Pat_id)
  If Pat_id <> 0 Then
   Call raZig.Open("SELECT * FROM `diagnosen` WHERE pat_id = " & Pat_id & " AND icd like '" & "F17" & "%" & "' AND diagsicherheit in (""G"",""V"",""Z"")  AND COALESCE(f6010,0)=0 ", DBCn, adOpenDynamic, adLockReadOnly)
   If Not raZig.EOF Then
    If InStrB(raZig!DiagText, "Z.n") <> 0 Or raZig!DiagSicherheit = "Z" Then
     If WieTabak <> früher And WieTabak <> vorlangem Then WieTabak = früher
    Else
     WieTabak = aktuell
    End If
   End If
  End If
End Function ' WieTabak

Function testvergleicheT()
 Dim rs As New ADODB.Recordset, AusS$, e1 As ZigSt, E2 As ZigSt, e1W$, e2W$, Zahl&
 Call Lese.ProgStart
 Open pVerz & "tabakst.txt" For Output As #331
 rs.Open "SELECT d.icd, d.diagsicherheit, a.* FROM `anamnesebogen` a LEFT JOIN `diagnosen` d ON a.pat_id = d.pat_id AND d.icd like 'F17%'", DBCn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  e1 = doTabakSt(rs!Pat_id)
  E2 = doTabakStAlt(rs!Tabak)
  If e1 = aktuell Then e1W = "Aktuell" Else If e1 = früher Then e1W = "Früher " Else If e1 = nie Then e1W = "Nie    " Else e1W = "VorLang"
  If E2 = aktuell Then e2W = "Aktuell" Else If E2 = früher Then e2W = "Früher " Else If E2 = nie Then e2W = "Nie    " Else e2W = "VorLang"
  If e1 <> E2 And Not (rs!tabakex = "j" And E2 = nie) And rs!Pat_id <> 51 And rs!Pat_id <> 602 And Not (e1 = früher And E2 = vorlangem And IsNumeric(rs!tabakbis)) And Not (e1 = nie And (E2 = früher Or E2 = aktuell)) And Not (e1 = vorlangem And E2 = früher) And Not (e1 = früher And E2 = vorlangem) Then
'  If e1 = früher AND e2 = vorlangem AND IsNumeric(rs!tabakbis) Then
'  If e1 = früher AND e2 = vorlangem Then
  If True Then
   AusS = e1W & " " & e2W & " " & Right$(Space(5) & rs!Pat_id, 5) & " " & Left(rs!Nachname & Space(15), 15)
   AusS = AusS & Left(rs!ICD & rs!DiagSicherheit & Space(7), 7) & " "
   AusS = AusS & Left(rs!Tabak & Space(70), 70) & " "
   AusS = AusS & Left(rs!tabakakt & Space(15), 10) & " "
   AusS = AusS & Left(rs!tabakex & Space(15), 10) & " "
   AusS = AusS & Left(rs!tabakbis & Space(15), 15) & " "
   AusS = AusS & Left(rs!tabakmenge & Space(15), 10) & " "
   Print #331, AusS
   Zahl = Zahl + 1
   End If
  End If
  rs.Move 1
 Loop
 Print #331, "Zahl: " & Zahl
 Close #331
 MsgBox "Fertig"
 Call Shell("notepad " & pVerz & "tabakst.txt", vbMaximizedFocus)
End Function ' testvergleicheT

 Function doTabakSt(Pat_id&) As ZigSt
  Dim rs As New ADODB.Recordset, tStr$
  On Error GoTo fehler
  Call Lese.ProgStart
  rs.Open "SELECT vorgestellt, tabakex,tabakakt,tabakbis,tabakmenge FROM `anamnesebogen` WHERE pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
  If rs!tabakakt = "j" Or InStrB(rs!tabakakt, "gelegentlich") <> 0 Or InStrB(rs!tabakakt, "ja") <> 0 Then
   doTabakSt = aktuell
  ElseIf IsNull(rs!tabakex) Then
nie:
   doTabakSt = nie
  ElseIf rs!tabakex = "-" Or rs!tabakex = "" Or rs!tabakex = "n" Or rs!tabakex = "?" Or Left$(rs!tabakex, 2) = "n " Or rs!tabakex = "nein" Or rs!tabakex = "Nein" Or rs!tabakex = "nie" Or rs!tabakex = "Nie" Or InStrB(rs!tabakex, "selber nicht") <> 0 Then
   GoTo nie
  Else
   doTabakSt = früher
   If Not IsNull(rs!tabakbis) Then
    tStr = LTrim$(rs!tabakbis)
    If Left(tStr, 3) = "bis" Then tStr = LTrim$(Mid$(tStr, 4))
    If Left(tStr, 3) = "ca." Then tStr = LTrim$(Mid$(tStr, 4))
    If InStrB(tStr, "-") <> 0 Then
     tStr = LTrim$(Mid$(tStr, InStr(tStr, "-") + 1))
    End If
    If IsDate(tStr) Then
     If Now() - CDate(tStr) > 12 * 365 Then
      doTabakSt = vorlangem
     End If
    Else
     If IsNumeric(Left$(tStr, 4)) Then tStr = Left$(tStr, 4)
     If IsNumeric(tStr) Then
      If Year(rs!Vorgestellt) - tStr >= 12 Then
       doTabakSt = vorlangem
      End If
     ElseIf tStr Like "*vor *" Then
      tStr = Mid(tStr, InStr(tStr, "vor ") + 4)
      If tStr Like "*J*" Then
       tStr = Left$(tStr, InStr(tStr, "J") - 1)
      ElseIf tStr Like "*a*" Then
       tStr = Left$(tStr, InStr(tStr, "a") - 1)
      End If
      If IsNumeric(tStr) Then
       If CDbl(tStr) >= 12 Then
        doTabakSt = vorlangem
       End If
      End If
     ElseIf InStrB(tStr, "80er J") <> 0 Then
      doTabakSt = vorlangem
     End If
    End If
   End If
  End If
      Exit Function
fehler:
     Dim AnwPfad$
    #If VBA6 Then
     AnwPfad = CurrentDb.name
    #Else
     AnwPfad = App.path
    #End If
    Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doTabakSt/" + AnwPfad)
     Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
     Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
     Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
    End Select
 End Function ' doTabakSt
 
 Function doTabakStAlt(Text) As ZigSt  ' TabakStatus "parsen"
  Dim KommaStelle%, FrüherStelle%, bisStelle%, MinusStelle%, LJStelle% ' 0 = keiner, 1 = früher, 2 = aktuell, 3 = vor mehr als 15 Jahren
  Dim jahrk$, obdatum%, Datu As Date, sp$()
  Dim rSn$, rSnpos&
  On Error GoTo fehler
  If Not IsNull(Text) Then
   Text = Trim$(Text)
   rSnpos = InStr(Text, "RauchenSienoch")
   If rSnpos > 0 Then
    rSn = Mid$(Text, rSnpos + 16)
    If InStrB(rSn, "RauchenSienoch") <> 0 Then rSn = Mid$(rSn, InStr(rSn, "RauchenSienoch") + 16)
   End If
   If Left(Text, 1) = "j" Then
    sp = Split(Text, "Frühergeraucht?")
    If Trim$(sp(0)) = ",wieviel?" Or obNein(Left(sp(0), 1)) Then
     doTabakStAlt = nie
    Else
     If rSnpos > 0 Then
      If obNein(rSn) Then doTabakStAlt = früher Else If (obNein(sp(0)) Or InStrB(rSn, "j") <> 0) Then doTabakStAlt = aktuell
     Else
      If LCase(Left(Text, 1)) = "j" Then
       If InStrB(Text, "is") <> 0 Then
        doTabakStAlt = früher
       Else
        doTabakStAlt = aktuell
       End If
      ElseIf obNein(sp(0)) Then
       doTabakStAlt = aktuell
      ElseIf UBound(sp) < 1 Then
       doTabakStAlt = nie
      ElseIf obNein(sp(1)) Then
       doTabakStAlt = nie
      Else
       doTabakStAlt = früher
      End If
     End If
    End If
   ElseIf Left(rSn, 1) = "j" Then 'obNein(rSn) = 0 Then 4.4.07 Pat. 1716
    doTabakStAlt = aktuell
   ElseIf Left(Text, 1) = "n" Or Text = "Frühergeraucht? n" Or Text = "Frühergeraucht? -" Or Left(Text, 1) = "," Then
    doTabakStAlt = nie
   Else
    KommaStelle = InStr(Text, ",")
    MinusStelle = InStr(Text, "-")
    bisStelle = InStr(Text, "bis")
    LJStelle = InStr(Text, "LJ")
    FrüherStelle = InStr(Text, "früher")
    obdatum = 0
    If MinusStelle = 1 Or bisStelle = 1 Then
     If MinusStelle = 1 Then jahrk = Mid$(Text, 2)
     If bisStelle = 1 Then jahrk = Mid$(Text, 4)
     jahrk = LTrim$(jahrk)
     If InStrB(jahrk, " ") <> 0 Then
      jahrk = Left(jahrk, InStr(jahrk, " ") - 1)
     End If
     If InStrB(jahrk, ".LJ") <> 0 Then
      jahrk = Left(jahrk, InStr(jahrk, ".LJ") - 1)
     End If
     If InStrB(jahrk, "/") <> 0 Then
      jahrk = Mid$(jahrk, InStr(jahrk, "/") + 1)
     End If
     If IsDate(jahrk) Then
      Datu = CDate(jahrk)
      obdatum = -1
     Else
      jahrk = Left(jahrk, 4)
     If IsNumeric(jahrk) Then
      obdatum = -1
     Else
      If IsNumeric(Left(jahrk, 2)) Then
       jahrk = Left(jahrk, 2)
       obdatum = -1
      End If
     End If
    End If
    
   End If ' MinusStelle = 1 Or bisStelle = 1 Then
   If Not IsNull(Text) And Text <> "" Then
    If bisStelle > 0 Or (FrüherStelle > 0 And (KommaStelle = 0 Or (KommaStelle > 0 And KommaStelle > FrüherStelle))) Then
     doTabakStAlt = früher
    ElseIf (FrüherStelle > 0 And KommaStelle > 0 And FrüherStelle > KommaStelle) Then
     doTabakStAlt = aktuell
    Else
     If InStrB(Text, "ja") <> 0 Or InStr(Text, "j") = 1 Then
      doTabakStAlt = aktuell
     Else
      If Text = "n" Or Text = "-" Or InStrB(Text, "nein") <> 0 Or InStrB(Text, "kein") <> 0 Or Text = "0" Then
       doTabakStAlt = nie
      Else
       If (InStrB(Text, "LJ") <> 0 And MinusStelle > 0) Or (MinusStelle = 1 And obdatum) Then
'        If Not obdatum Or (obdatum AND datu > Now + 15 * 365) Then
' muß noch getestet werden
         doTabakStAlt = früher
       Else
        doTabakStAlt = aktuell
       End If
      End If
     End If
    End If
   End If
  End If ' not ISNULL(Text
  End If
   If doTabakStAlt = früher And InStrB(Text, "biswann:") <> 0 Then
    bisStelle = InStr(Text, "biswann:") + Len("biswann:")
    jahrk = LTrim$(Trim$(Mid$(Text, bisStelle + 1)))
    obdatum = -1
   End If
  If doTabakStAlt = 1 And obdatum Then
   If IsDate(jahrk) Then
    Datu = CDate(jahrk)
   Else
    Dim mn$
    mn = MachNumerisch(jahrk)
    If IsDate("1.1." + mn) Then
     Datu = CDate("1.1." + mn)
    ElseIf IsDate(mn) Then
     Datu = CDate(mn)
    End If
   End If
   If Datu < Now - 15 * 365 Then
    doTabakStAlt = vorlangem
   End If
  End If
    
  Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doTabakStAlt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
 End Function ' doTabakStAlt

Function ZSU$(Quelle$)
Dim i%, ZS$, bq$, bz$
 On Error GoTo fehler
ZS = vNS
For i = 1 To Len(Quelle)
 bq = Mid$(Quelle, i, 1)
 Select Case bq
  Case "ß"
   bz = "á"
  Case "ä"
   bz = "„"
  Case "ö"
   bz = "”"
  Case "ü"
   bz = ""
  Case "Ä"
   bz = "Ž"
  Case "Ö"
   bz = "™"
  Case "Ü"
   bz = "š"
  Case "á"
   bz = " "
  Case "à"
   bz = "…"
  Case "é"
   bz = "‚"
  Case "è"
   bz = "Š"
  Case "í"
   bz = "¡"
  Case "ì"
   bz = ""
  Case "ó"
   bz = "¢"
  Case "ò"
   bz = "•"
  Case "ú"
   bz = "£"
  Case "ù"
   bz = "—"
  Case "â"
   bz = "ƒ"
  Case "ê"
   bz = "ˆ"
  Case "î"
   bz = "Œ"
  Case "ô"
   bz = "“"
  Case "û"
   bz = "–"
  Case "µ"
   bz = "æ"
  Case "²"
   bz = "ý"
  Case "³"
   bz = "ü"
  Case "§"
   bz = "õ"
  Case "€"
   bz = "_"
  Case Else
   bz = bq
 End Select
 ZS = ZS + bz
Next i
ZSU = ZS
If ZSU <> Quelle Then
 Debug.Print "ZSU: " + ZSU + " -> " + Quelle
End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZSU/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ZSU

Public Sub inTMAnz(Pat_id&)
On Error GoTo fehler
#Const defAutoIt = False
#If defAutoIt Then
    Dim cAutoit As New AutoItX3Lib.AutoItX3
    Dim Verz$, erg%
    erg = cAutoit.WinExists("TurboMed")
    If erg = 0 Then
     Verz = cAutoit.RegRead("HKEY_LOCAL_MACHINE\Software\TurboMed EDV GmbH\TurboMed\Current", "RegisterPath")
    ' MsgBox(0,"TM-Pfad",$Verz)
     cAutoit.rUn (Verz & "turbomed.exe")
     Call cAutoit.WinWait("TurboMed - ", "")
     If Not cAutoit.WinActive("TurboMed - ", "") Then Call cAutoit.WinActivate("TurboMed - ", "")
     Call cAutoit.WinWaitActive("TurboMed - ", "")
     Call cAutoit.send("+{TAB}gerald{tab}mond!a")
    End If
    Call cAutoit.WinWait("TurboMed", "")
    If Not cAutoit.WinActive("TurboMed", "") Then Call cAutoit.WinActivate("TurboMed", "")
    Call cAutoit.WinWaitActive("TurboMed", "")
    Call cAutoit.send("{F12} p")
    Call cAutoit.WinWait("TurboMed", "")
    If Not cAutoit.WinActive("TurboMed", "") Then Call cAutoit.WinActivate("TurboMed", "")
    Call cAutoit.WinWaitActive("TurboMed", "")
    Dim pstr$, i%
    If False Then
     pstr = Pat_id 'adoRS!Pat_id
     For i = 1 To Len(pstr)
      Call cAutoit.send(" ")
      Call cAutoit.send(Mid(pstr, i, 1))
    ' Call cAutoit.sleep(100)
     Next i
     Debug.Print "adors!pat_id: " & pstr
    End If
    'Exit Sub
    Call cAutoit.WinWait("TurboMed", "")
    If Not cAutoit.WinActive("TurboMed", "") Then Call cAutoit.WinActivate("TurboMed", "")
    Call cAutoit.WinWaitActive("TurboMed", "")
    Call cAutoit.send("{ENTER}")
    Call cAutoit.WinWait("TurboMed", "")
    If Not cAutoit.WinActive("TurboMed", "") Then Call cAutoit.WinActivate("TurboMed", "")
    Call cAutoit.WinWaitActive("TurboMed", "")
    Call cAutoit.send("{F3}")
    Call cAutoit.WinWait("TurboMed", "")
    If Not cAutoit.WinActive("TurboMed", "") Then Call cAutoit.WinActivate("TurboMed", "")
    Call cAutoit.WinWaitActive("TurboMed", "")
    Call cAutoit.send("{F11}")
#Else
 Dim hnd&
 hnd = FensterHandle("TurboMed.exe")
 If hnd <> 0 Then
  On Error Resume Next
  AppActivate "TURBOMED", True
  If Err.Number <> 0 Then Exit Sub
  Pause (Pausenlänge)
  SendKeys "{ESC}", True
  Pause (Pausenlänge)
  SendKeys "{ESC}", True
  Pause (Pausenlänge)
  SendKeys "{ESC}", True
  Pause (Pausenlänge)
  SendKeys "{ENTER}", True
  Pause (Pausenlänge)
  SendKeys "{F12}", True
  Pause (Pausenlänge)
  SendKeys "p", True
  Pause (Pausenlänge)
  SendKeys "{bs}" & Pat_id & "", True
  Pause (Pausenlänge)
  SendKeys "{ENTER}", True
  Pause (Pausenlänge)
  SendKeys "{F3}", True
  Pause (Pausenlänge)
  SendKeys "%{F1}", True
  On Error GoTo fehler
'  Pause (Pausenlänge)
'  AppActivate "TurboMed", True
 End If
#End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in hollabor/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' InTMAnz(Pat_id&)

Public Function GetRandomNum&(Min&, MAX&)
    GetRandomNum = Int(Rnd * (MAX - Min + 1) + Min)
End Function ' GetRandomNum

'Public Function s2test(Pat_id&)
'Lese.ProgStart
'Dim z&
'Dim r2 AS New ADODB.Recordset
'   r2.Open "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_id & " AND abkü RLIKE '^hba[1c]' AND CAST(wert AS decimal) < 22 UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_id & " AND abkü RLIKE '^hba[1c]' AND CAST(wert AS decimal) < 22) i GROUP BY pat_id,zeitpunkt,abkü,wert ORDER BY zeitpunkt ", DBCn, adOpenStatic, adLockReadOnly '
'   Do While Not r2.EOF
'    z = z + 1
'    r2.MoveNext
'   Loop
'   Debug.Print z
'   z = 100
'   Set r2 = hollabor(Pat_id, "HBA[1C]", 0, 22, 0, z)
'Debug.Print z
'End Function
' zahl liefert die Zahl der Datensätze zurück, wenn nicht -1 übergeben wird
Public Function hollabor(Optional PatID& = 0, Optional Abkü$ = "", Optional zpkl As Date = 0, Optional wertkl& = 0, Optional obnachgruppe% = 0, Optional ByRef Zahl& = -1, Optional Einheit$ = "", Optional gz& = -1, Optional dzz& = -1, Optional dSL As SortierListe) As ADODB.Recordset
 Static rs As New ADODB.Recordset, sql$
 Static altPID&, altzahl&, altab$, altzpkl#, altwk&, altobng%, alteh$, altgz&, altdz&
 Dim SD As SortierDatum
 Static hdSL As New SortierListe ' dzz = datumszahl, wenn nicht nach Gruppen gruppiert (also nach Zeitpunkt sortiert)
 Dim vgru$ ' vorherige Gruppe
 On Error GoTo fehler
 If PatID = altPID And altab = Abkü And altzpkl = zpkl And altwk = wertkl And altobng = obnachgruppe And alteh = Einheit Then
  If Not rs.BOF Then rs.MoveFirst
 Else
  altzahl = 0
  If PatID <> altPID Then
   altgz = 0
   altdz = 0
   Set hdSL = Nothing
  End If
  Set rs = Nothing
'#Const problematisch = True
#If problematisch Then
  'DBCn.Execute "SET @patid=" & CStr(PatID)
  'DBCn.Execute "SET @abkue='" & Abkü & "'"
  'DBCn.Execute "SET @einheit='" & Einheit & "'"
  'sql = "SET @zpkl="
  sql = ""
  If zpkl = 0 Then sql = sql + "0" Else sql = sql + DatFor_k(zpkl)
  'DBCn.Execute sql
  DBCn.Execute "SET @patid=" & CStr(PatID) & ",@abkue='" & Abkü & "',@einheit='" & Einheit & "',@zpkl=" & sql & ",@wertkl=" & wertkl
' rs.CursorLocation = adUseServer
  sql = "SELECT * FROM geslab" + IIf(obnachgruppe, " WHERE (reihe <> 999 OR ISNULL(reihe)) GROUP BY gruppe, reihe, abkü, einheit,ung,ong ORDER BY gruppe,reihe", " ORDER BY zeitpunkt DESC")
  rs.Open sql, DBCn, adOpenDynamic, adLockReadOnly
#Else
 Dim par$
 ' 4.7.20: wert statt einheit eingesetzt, da in labor1 und labor2 offenbar verschiedene Einheiten verwendet werden, z.B. ml/min = ml/mn/1.73 m²
 par = IIf(obnachgruppe, " WHERE (reihe <> 999 OR ISNULL(reihe)) GROUP BY gruppe, reihe, abkü, einheit,ung,ong ORDER BY gruppe,reihe", IIf(Abkü <> "", "WHERE abkü=""" & Abkü & """" & IIf(Einheit <> "", " AND einheit =""" & Einheit & """", "") & IIf(zpkl <> 0, " AND zeitpunkt<" & Format(zpkl, "yyyymmdd"), ""), "") & " GROUP BY zeitpunkt DESC,abkü,wert ORDER BY zeitpunkt DESC")
 rs.Open "call geslabdp(" & CStr(PatID) & ",'" & par & "')", DBCn, adOpenStatic, adLockReadOnly
#End If
 End If ' altpid<>pid
#If langsamer Then
 If Zahl <> -1 Then Zahl = DBCn.Execute("SELECT COUNT(0) Zahl FROM (" & sql & ")")!Zahl
#Else
' das folgende geht bei ca. 1500 Datensätzen in 0,05s statt mit COUNT(0) in 0,15s
 If Zahl <> -1 Or gz <> -1 Or dzz <> -1 Then ' wenn Zahl gewünscht wird
  If (Zahl <> -1 And altzahl = 0) Or (gz <> -1 And altgz = 0) Or (dzz <> -1 And altdz = 0 And Not obnachgruppe) Then ' ... aber nicht vorliegt
   altgz = 0
   If Not obnachgruppe Then
    altdz = 0
    Set hdSL = Nothing
   End If
   Do While Not rs.EOF ' ... dann ausrechnen
    If rs!gruppe <> vgru Then
     altgz = altgz + 1
     vgru = rs!gruppe
    End If
    ' wenn der Patient einmal ohne obnachgruppe und dann mit aufgerufen wird, dann kann wg. Performance altdz vom vorigen Aufruf jetzt zurückgeliefert werden
    If Not obnachgruppe Then
     Set SD = New SortierDatum
     SD.Datum = Int(rs!Zeitpunkt)
     If hdSL.GetItem(SD) Is Nothing Then
      hdSL.sCAdd SD
      altdz = altdz + 1
     End If
    End If ' not obnachgruppe
    altzahl = altzahl + 1
    rs.MoveNext
   Loop
   If Not rs.BOF Then rs.MoveFirst
  End If ' ... dann
  If Zahl <> -1 Then Zahl = altzahl ' ... liefern
  If gz <> -1 Then gz = altgz
  If dzz <> -1 Then
   dzz = altdz
   Set dSL = hdSL
  End If
 End If ' Zahl <> -1
#End If
 altPID = PatID
 altab = Abkü
 altzpkl = zpkl
 altwk = wertkl
 altobng = obnachgruppe
 alteh = Einheit
 Set hollabor = rs
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in hollabor/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' hollabor

Public Function testlab(pid&)
 Dim rs As New ADODB.Recordset
 Lese.ProgStart
' DBCn.Execute ("SET @patid=" & pid)
' rs.Open "SELECT * FROM geslab", DBCn, adOpenStatic, adLockReadOnly
Const obnachgruppe% = 1
Dim par$
par = IIf(obnachgruppe, " WHERE (reihe <> 999 OR ISNULL(reihe)) GROUP BY gruppe, reihe, abkü, einheit,ung,ong ORDER BY gruppe,reihe", " GROUP BY zeitpunkt DESC,abkü,einheit ORDER BY zeitpunkt DESC")
rs.Open "call geslabdp(" & pid & ",'" & par & "')", DBCn, adOpenStatic, adLockReadOnly
 TabAusgeb rs, Lese, True, , , , , , "Labor " & pid
End Function ' testlab(pid&)

'#If False Then
Public Function testhl()
' für HAAkt z.B. geht das nicht, da es kein Lese enthält
' Call Lese.ProgStart
 Dim rs As ADODB.Recordset
 Set rs = hollabor(14, "")
' für HAAkt z.B. geht das nicht, da es kein Lese enthält
' TabAusgeb rs, Lese, True, , , , , , "Labor 14"
 rs.Close
' Set rs = hollabor(2, "Ery")
' TabAusgeb rs, Lese, True, , , , , , "Labor 2"
End Function
'#End If


