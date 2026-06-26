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
Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy&)

Const BreiteSp1% = 28
Const LEDatei$ = "Leistungsexport_Porto.txt"

Const HADBName$ = "haerzte"
Public Const vgbVerspätung% = 21 ' wie lange ins neue Quartal die alte Abrechnung geht
Public Verspätung$
'Public Const kvneu = "SELECT GROUP_CONCAT(DISTINCT nachname) haname, ort, CAST(LEFT(bsnr,7) AS char) kvnu, CONCAT(LEFT(bsnr,2),'/',MID(bsnr,3,5)) kvnr, REPLACE(tel1.tel,'-','') tel1, REPLACE(fax,'-','') fax1, IF(obweibl,'Frau','Herr') anrede, titel, vorname, nachname, " & _
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
'Public Const kvneu1 = "SELECT GROUP_CONCAT(DISTINCT nachname) haname, ort, CAST(LEFT(bsnr,7) AS char) kvnu, " & _
                      "if(ISNULL((SELECT genehmigung FROM " & hadbname & ".arzt_has_genehmigung ahg1 LEFT JOIN " & hadbname & ".genehmigung g1 ON g1.idgenehmigung = ahg1.genehmigung_id WHERE ahg1.arzt_id = idarzt AND genehmigung = 'DMP-DM1_Koordinierender Arzt_Hausarzt')),'0','1') dmpt1, " & _
                      "if(ISNULL((SELECT genehmigung FROM " & hadbname & ".arzt_has_genehmigung ahg2 LEFT JOIN " & hadbname & ".genehmigung g2 ON g2.idgenehmigung = ahg2.genehmigung_id WHERE ahg2.arzt_id = idarzt AND genehmigung = 'DMP-DM2_Koordinierender Arzt')),'0','1') dmpt2 " & _
                      "FROM " & hadbname & ".bs " & _
                      "LEFT JOIN " & hadbname & ".ort ON ort_id = idort " & _
                      "LEFT JOIN " & hadbname & ".arzt_has_bs ahb ON ahb.bs_id = idbs " & _
                      "LEFT JOIN " & hadbname & ".arzt a ON a.idarzt = ahb.arzt_id " & _
                      "GROUP BY kvnu"
                      
' 11.9.15: kommt offenbar nicht vor, müsste ggf. noch myEFrag "SET GROUP_CONCAT_MAX_LEN = 70" vorangestellt werden
Public Const kvneu2 = "SELECT GROUP_CONCAT(DISTINCT nachname) haname, ort, CAST(LEFT(bsnr,7) AS char) kvnu, CONCAT(LEFT(bsnr,2),'/',MID(bsnr,3,5)) kvnr, REPLACE(tel1.tel,'-','') tel1, REPLACE(tel2.tel,'-','') tel2, REPLACE(fax1.fax,'-','') fax1, REPLACE(fax2.fax,'-','') fax2, mail1.mail email, IF(obweibl,'Frau','Herr') anrede, titel, vorname, nachname, " & _
                      "MAX(IF(ISNULL((SELECT genehmigung FROM " & HADBName & ".arzt_has_genehmigung ahg1 LEFT JOIN " & HADBName & ".genehmigung g1 ON g1.idgenehmigung = ahg1.genehmigung_id WHERE ahg1.arzt_id = idarzt AND genehmigung = 'DMP-DM1_Koordinierender Arzt_Hausarzt')),'0','1')) dmpt1, " & _
                      "MAX(IF(ISNULL((SELECT genehmigung FROM " & HADBName & ".arzt_has_genehmigung ahg2 LEFT JOIN " & HADBName & ".genehmigung g2 ON g2.idgenehmigung = ahg2.genehmigung_id WHERE ahg2.arzt_id = idarzt AND genehmigung = 'DMP-DM2_Koordinierender Arzt')),'0','1')) dmpt2, " & _
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

Public Const artSpezÄrzte$ = "'ni','gstel','gs','rz','ep','bga','tk','APK','wd','ah','ta','tb','tt','tn'"
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

Public Const artSpezBerat$ = artSpezÄrzte & ",'wr','jl','ga','ih','cr','tn','be','lf','kb','mn','lt','sk'"
' wr = Walburga Roßmeier
' jl = Juliane Lange
' ga = Greß Angelika
' ih = Ines Hoffmann
' cr = Cornelia Reindl
' tn = Tanja Nuber
' be = Bender Elena
' lf = Larissa Fuchs
' kb = Katja Butz
' mn = Milande Nana
' lt = Luisa Tepe
' sk = Sarah Kreis

Public Const artSpezMA$ = "'tst','ke','hz','ns','mh','ag','ph','pq','er','ds','st','eb','us','sn','vb','mip','mm','rc','ik','ks','sb','cb','sp','ir','as','sa','sta','eg','ans','mc','rb','mi','gr','bs','sf','fs','eo','cd','mk','nb','sas','mf','bt','ab','sh','an','lo','fm'"
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
' nb = Nina Birgmeir
' sas = Sabrina Strasser
' mf = Melina Fischhold
' bt = Brigitte Tziutzukis
' ab = Ayten Bülbül
' sh = Sibora Hasani
' an = Amssatou Ndjamawe
' tb = Texteintrag Befund
' lo = Laima Ostmeyer
' fm = Fetije Morina

' was im Arztbrief unter Verlaufsnotizen stehen soll
Public Const artVerlauf$ = "'notiz','mbf','telef'," & artSpezBerat & "," & artSpezMA & _
",'tb','th','tn','pa','beruf','sem','seminar','kind','haut','kard','aug','auge','augen'" & _
",'mess','ictauf','auto','insgd','vkg','vkgd','vkgd2'" & _
",'impf','imp','gyn','caro','carotis'" & _
",'schula','ass','kra','proc','aufgd','HbA1cfre'" & _
",'ARCHIE2','ARCHIE','PATNRALPHA','arch','FDIAB1','FDIAB2','BLANKO','NV','M-HI','REZ-K','REZ-X','EDIAB1','LAS','EDIAB2','lab','Novo Rapid','PBA','PRAXG','FT','REZ-G','N','KK','REZ-P','anam','EKG01','KH','HV13','M-HM','L','IE','HÄK'" & _
",'tv','schul','schulung','50g','dak','dakap','dakluts','daknp','dmperg','unin','c19i','cia','fana','fth','rauch'"

' Einträge, die nicht automatisch mit einer Organuntersuchung verbunden sind
Public Const artSpezEintr$ = artVerlauf & ",'beweg','bewegung','bew','bewg','bewe','bz','bztp','bzm'," & _
"'uzu','uz','hypo','GPD','GDP','ba','bla','kva','LZ','rp','re','B-rp','mu','mus','muster','au'," & _
"'fa','fam','familie','fams','rauch','rauchen','alko','pros','prost','prostata','uro','colo','coloauf'"
'Public Const artSpezEintr$ = "'notiz','mbf','telef'," & artSpezBerat & "," & artSpezMA & _
",'pa','fa','fam','familie','beruf','sem','seminar','kind','bz','rp','B-rp','mus','muster','uzu','uz','hypo','colo','coloauf','haut','kard','aug','auge','augen'" & _
",'beweg','bewegung','bew','bewg','bewe','mess','ictauf','auto','insgd','vkgd','vkgd2'" & _
",'pros','prost','prostata','uro','impf','imp','gyn','caro','carotis'" & _
",'mu','rauch','rauchen','alko','fams','schula','ass','kra','proc','au','GPD','GDP','ba','bla','kva','aufgd','HbA1cfre'" & _
",'ARCHIE2','ARCHIE','PATNRALPHA','arch','LZ','FDIAB1','FDIAB2','BLANKO','NV','M-HI','REZ-K','REZ-X','EDIAB1','LAS','EDIAB2','lab','Novo Rapid','PBA','PRAXG','FT','REZ-G','N','KK','REZ-P','anam','EKG01','KH','HV13','M-HM','L','IE','HÄK'" & _
",'tv','re','schul','schulung','50g','dak','dakap','dakluts','daknp','dmperg','unin','c19i','cia','fana','fth'"
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
' fana = Anamnese bei KIM-ArztBriefen

' Einträge mit Untersuchungen, die vorne im Patientenlaufzettel stehen
Public Const artSpezUS0$ = "'gewicht','gewi','gew','gw','rrvgl','rrvgln','bzvgl','bzvgln','fuß','füße','taille','urin'"
' gewicht, gewi, gew = Gewicht
' rrvgl = RR-Vergleich
' rrvgln = RR-Vergleich nicht möglich, Gerät nicht dabei
' bzvgl = Blutzuckervergleich
' bzvgln = Blutzuckervergleich nicht möglich, Gerät nicht dabei
' fuß,füße = genaue Fußinspektion
' taille = Taillenumfang
' urin = Urin-Teststreifenergebnis

' Einträge mit Untersuchungen, die nicht vorne im Patientenlaufzettel stehen
Public Const artSpezUS1$ = "'bzm','bztp','bks','anal','andm','andm2','angd','angd2','usal','usd','usdm','usdm1','usdm2','bauch','ufrag','uabfrag'" & _
",'cgma','doppler','dop','duplex','dup','belastun'" & _
",'sono','sd','UKG','Größe','groe','HbA1c','hyper','keto','wv','vw','kv','komp','kompr','debr','ulc','ulcus','EKG','LZRR','Lufu','gpt'" & _
",'lactoset','trop','temp','oGTT','bmi','hüfte','rr','puls','GDT','bef','OAU','OA','inj','inf','ht','htt','ADL','TUG','247','284','bar','who','vac','vacc','grippe','fbef','fbild'"
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
' dup = Duplexbefund
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
' fbef = über KIM importierte Befundzeile
' fbild = über KIM importiertes Bild

' Einträge, die mit einer Organuntersuchung verbunden sind
Public Const artSpezUS$ = artSpezUS0 & "," & artSpezUS1

' für Liste 112
Public Const artSpezSonst$ = "'PKon','kva','rcg','dd','dg','ddg','bddg','rdg','lar','plar','spbrp','htxt','txt','priv','bez','rech'" & _
",'pvs','mhng','Ch.B.','prp','unbek','WZeit','ict','ct','oad','typ1','BfA','PGeb','haem','Text','kgä','kgp','ksä','ksp','kmä','kmp'" & _
",'mimp','uebw','abstr','cib','nfd','dpe','mpa'"
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
' mpa = Medikamentenplan aktualisert

Public Const artspezG$ = artSpezEintr & "," & artSpezUS
Public Const AuffArtSql = vbCrLf & _
 "SELECT art, Pat_ID, gesname(pat_id) Name, Zeitpunkt, Ersteller, Inhalt FROM eintraege e FORCE INDEX (art)" & vbCrLf & _
 "WHERE NOT art IN (" & artspezG & "," & artSpezSonst & vbCrLf & _
 ",'pdf','utxt','UEBLABOR','med','ts','ana','RRS','RRD','mp','HEILMITTEL','tm','EARZT','ANFLABOR','HMPHYSIK','dmph','covze','iplan','DIAGTXT','link','tf','ABS','ti','EDMPKHK','REHASPORT','ssd','usd','BTM','sob','EDMPAB','tc','son','ATT','DIABEM','BRIEF','AUWEITER','AWHB'," & vbCrLf & _
 "'BILD','COVGE','DOKUAB','dup','EDMPCOPD','gluv','HKPFLEGE','HMERGO','KRBEFOERD','rrv','BSCH','EXTERN','foto','jc','vk','vp')" & vbCrLf & _
 "AND NOT art RLIKE '^[0-9]*$|^eDMPDM|^DMPDTYP|Dokumentation Diabetes|XCA|XCHOL|XCREA|XCRP|XERY|XFERRI|XFT4|XGFR|XGGT|XGLU|XGOT|XGPT|XHB|XHBA1C|XHDL|XHKT|XHSRE|XHST|XINR|XK|XLDH|XLDL|XLEUKO|XMCH|XMCV|XNA|XNT-PROBNP|XQUICK|XTHRO|XTRI|XTSH|XVITB12|XVITD3|XXDIAB'" & vbCrLf & _
 "AND NOT (art LIKE 'X%' AND ersteller IN ('hm','gs'))" & vbCrLf & _
 "AND zeitpunkt > 20231231 " & vbCrLf & _
 "ORDER BY zeitpunkt<20250325,art,zeitpunkt DESC; "
'Public Const artinartspez$ = "(art IN " & artSpez
Public sql$, sql1$ ' SQL-Text für alle möglichen Abfragen
Dim QMdbAkt$, nzw$
Dim DmPStrS As New CString
Dim Pat_ID& ' für dii(
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

Public Enum DMPartEnum
 [_First] = -1
 DMPDm
 DMPKhk
 DMPAb
 DMPCopd
 DMPBK
 DMPOsteo
 DMPRa
 DMPChi
 DMPRS
 DMPAd
 DMPDep
 [_Last]
End Enum

Public Enum DMPEnum
  unb = 0
  nein
  HA
  hier
  ausg
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
 Pat_ID As Long
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
 geschlecht As String
 Anschrzus As String
 Postleitzahl As String
 KVBereich As String * 2
 PrivatTel As String
 KVKStatus As String * 1
 bPerG As String * 2
 DMPKnZ As String * 2
 IKPrae As String * 2
 IK As String * 9 '7 => lieber 9 s. Wiki
 IKName As String
 VschBeg As Date
 VschEnd As Date
 lVorl As Date
 ausgst As Date
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
 SW(5) As String ' As Date ' Schulungen wahrgenommen ' (0) = D.m., (1) = Hypertonie, (2) = keine
 Ds(2) As Date ' DiabSchulung
 HS(2) As Date ' HyperSchulung
 VorM(13) As Integer
 obmetf As AntidiabMedType
 obGlib As AntidiabMedType
 obGlucI As AntidiabMedType
 obSHGlin As AntidiabMedType
 obGlit As AntidiabMedType
 obDpp4 As AntidiabMedType
 obGlp1 As AntidiabMedType
 obSglt2 As AntidiabMedType
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
 kGR As Integer ' Körpergröße
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
 ab315 As Integer
 ab317 As Integer
 ab1023 As Integer
 x_GesName As String
 x_DoklangName As String
 x_DokName As String
 x_DokuDat As Date
 x_DmTyp As String
 x_Bbk As String
 x_Gewi As String
 x_DMPHbA1c As String
 x_DMPeGFR As String
 x_DMPUAlb As String
 x_Spät As String
 x_Inj As String
 x_Pulsstatus As String
 x_SensText As String
 x_wrstr As String
 x_Ulcus As String
 x_Wundinfektion As String
 x_Intervall As String
 x_Ereig As String
 x_HypoStr As String
 x_Insulin As String
 x_Metformin As String
 x_Sonstige As String
 x_Sglt2 As String
 x_Glp1 As String
 x_Thromb As String
 x_OAK As String
 x_SchulEmpf As String
 x_SchulWahrg As String
 x_HSchulWahr As String
 x_SchulStr As String
 x_InfoAng As String
 x_HZiel As String
 x_BhdF As String
 x_BhdD As String
 x_Aug As String
 x_EmpfItv As String
End Type ' DMPClass

Enum LabArt
 LabArt0
 LA_HbA1c
 LA_HbA1cmm
 LA_Krea
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
 LA_GOT
 LA_THR
 LA_AlbS
End Enum ' LabArt
' statische Variable für die Funktion LabPat
Public alt_la As LabArt, altPID&, aktlwx&, lwZahl&

Public Type labtyp
 Abkü As String
 Zp As Date
 WertSg As String ' Wert-String, mit "," als Dezimaltrennzeichen;
 Einheit As String
End Type
 
Public lab() As labtyp

'#Const sqllangsam = True
#If sqllangsam Then
' nach Vorbild von Public FUNCTION LabPat(la AS LabArt, Pat_id&) AS labtyp
Public Function LetztLab(pid&, Abkü$, Einh$, Zp As Date) As labtyp
 Dim gefunden%, i&
 On Error GoTo fehler
 ' bei neuem Patienten sein ganzes Labor einlesen
 If pid <> altPID Then
  Dim rsl As New ADODB.Recordset ' Labor
  Set rsl = hollabor(PatID:=pid, Abkü:="", zpkl:=0, wertkl:=0, obnachgruppe:=0, Zahl:=lwZahl)
  If Not rsl.BOF And lwZahl Then
   ReDim lab(lwZahl)
   i = 0
   Do While Not rsl.EOF
    lab(i).Abkü = rsl!Abkü
    lab(i).WertSg = REPLACE$(rsl!Wert, ".", ",")
    lab(i).Einheit = rsl!Einheit
    lab(i).Zp = rsl!Zeitpunkt
    i = i + 1
    rsl.MoveNext
   Loop
  End If
  Set rsl = Nothing
 End If ' pid <> altPID THEN
 gefunden = 0
 If lwZahl Then
  For i = 0 To lwZahl
   If lab(i).Abkü = Abkü And lab(i).Einheit = Einh And lab(i).Zp < Zp Then
    LetztLab.Abkü = lab(i).Abkü
    LetztLab.Einheit = lab(i).Einheit
    LetztLab.WertSg = REPLACE$(lab(i).WertSg, ".", ",")
    LetztLab.Zp = lab(i).Zp
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LetztLab/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LetztLab(
#End If


' wenn nichts (mehr) gefunden, dann wird zp auf 0 gesetzt
Public Function LabPat(LA As LabArt, pid&, Optional naechster%) As labtyp
 Dim gefunden%
 On Error GoTo fehler
 ' bei neuem Patienten sein ganzes Labor einlesen
 If pid <> altPID Then
  Dim rsl As New ADODB.Recordset ' Labor
  Dim i&
  Set rsl = hollabor(pid, "", 0, 0, 0, lwZahl)
  If Not rsl.BOF And lwZahl Then
   ReDim lab(lwZahl)
   i = 0
   Do While Not rsl.EOF
    lab(i).Abkü = rsl!Abkü
    lab(i).WertSg = REPLACE$(rsl!Wert, ".", ",")
    lab(i).Einheit = rsl!Einheit
    lab(i).Zp = rsl!Zeitpunkt
    i = i + 1
    rsl.MoveNext
   Loop
  End If
  Set rsl = Nothing
 End If
 ' dann entscheiden, ob mit dem Satzzeiger von vorne angefangen werden muss
 Dim angefx&
 If pid <> altPID Or alt_la <> LA Then
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
   If lab(aktlwx).WertSg <> "" Then
    If obLabI(LA, lab(aktlwx)) Then
     gefunden = True
     Exit Do
    End If
   End If
   angefx = True
  Loop
  If gefunden Then
   LabPat.Abkü = lab(aktlwx).Abkü
   LabPat.Einheit = lab(aktlwx).Einheit
   LabPat.WertSg = REPLACE$(lab(aktlwx).WertSg, ".", ",")
   LabPat.Zp = lab(aktlwx).Zp
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
 alt_la = LA
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabPat/" + AnwPfad)
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

'Function LabEPat(la AS LabArt, Pat_id&) As Adodb.Recordset
' SET LabEPat = New ADODB.Recordset
' myFrag LabEPat, LabEPatS(la, Pat_id).Value
'End FUNCTION ' LabEPat

Function obLabI%(LA As LabArt, lT As labtyp) ' ob Labor interessant
 Select Case LA
  Case LA_HbA1c:
   obLabI = ((lT.Abkü Like "HBA[C1]*" Or lT.Abkü = "XHBA1C") And MachNumerisch(lT.WertSg) < 22) ' 30.8.14: wegen falsch eingetragener Fremdlabore gestrichen: AND `einheit` = '%'
            ' identisch, wg.Kompatibilität nicht genommen: "RLIKE 'hba[c1]' AND einheit = '%'"
  Case LA_HbA1cmm:
   obLabI = ((UCase$(lT.Abkü) Like "HBA[C1]*" Or lT.Abkü = "XHBA1C") And MachNumerisch(lT.WertSg) >= 22) ' 30.8.14: wegen falsch eingetragener Fremdlabore gestrichen: AND `einheit` = '%'
  Case LA_Krea:
  ' abkü IN ('creat','krea02','krea','krea02','kres') AND einheit='mg/dl'
   obLabI = ((lT.Abkü = "CREAT" Or lT.Abkü = "KRE02" Or lT.Abkü = "KREA" Or lT.Abkü = "KREA02" Or lT.Abkü = "KRES" Or lT.Abkü = "XCREA") And lT.Einheit = "mg/dl")
  Case LA_AlbCre:
  ' zu koordinieren mit doViewserstellen: obnephrop und labpath.cpp pVirtfuehraus
   obLabI = ((lT.Abkü = "ALBCRE" Or lT.Abkü = "ALBKRE" Or lT.Abkü = "ALBQ" Or lT.Abkü = "ALBUM" Or lT.Abkü = "ALBUP") And (lT.Einheit Like "mg/g *" Or lT.Einheit = "keine Einheit" Or lT.Einheit = "" Or lT.Einheit = "kA" Or lT.Einheit = "'kA'")) Or ((lT.Abkü = "ALBU" Or lT.Abkü = "ALBUMU") And lT.Einheit = "mg/l")
  Case LA_Kali:
   obLabI = (lT.Abkü = "K" Or lT.Abkü = "KALI" Or lT.Abkü = "XK")
  Case LA_Chol:
   obLabI = (lT.Abkü = "CHOL" Or lT.Abkü = "XCHOL")
  Case LA_Hdl:
   obLabI = (lT.Abkü = "HDL" Or lT.Abkü = "HDLC" Or lT.Abkü = "XHDL")
  Case LA_LDL:
   obLabI = ((lT.Abkü Like "LDL*" Or lT.Abkü = "XLDL") And lT.Einheit = "mg/dl")
  Case LA_eGFR:
   obLabI = (lT.Abkü Like "GFR*" Or UCase$(lT.Abkü) Like "GFC*" Or lT.Abkü Like "CREACL*" Or lT.Abkü = "MDRD" Or lT.Abkü = "XGFR") ' GFCK4
   ' "GFR','GFRCYS','GFRK','GFRM','GFRT','GFRW','CREACL','GFC','GFCK1','GFCK2','GFCK3','GFCK4','GFCM1','GFCW1','GFCW2','MDRD"
  Case LA_TG:
   obLabI = (lT.Abkü = "TRI" Or lT.Abkü = "TRIG" Or lT.Abkü = "NTFE" Or lT.Abkü = "TRIGLK" Or lT.Abkü = "XTRI")
  Case LA_GGT:
   obLabI = (lT.Abkü = "GAGT" Or lT.Abkü = "GAGT02" Or lT.Abkü = "GGT" Or lT.Abkü = "GT" Or lT.Abkü = "YGT" Or lT.Abkü = "XGGT")
  Case LA_GPT:
   obLabI = (lT.Abkü = "GPT" Or lT.Abkü = "GPTR02" Or lT.Abkü = "GPTR03" Or lT.Abkü = "PT" Or lT.Abkü = "XGPT") ' prognostiziert
  Case LA_FERR:
   obLabI = (lT.Abkü = "FERR" Or lT.Abkü = "FERR01" Or lT.Abkü = "FERRI" Or lT.Abkü = "FERRIT" Or lT.Abkü = "FERRO" Or lT.Abkü = "XFERRI")
  Case LA_DIGO:
   obLabI = (lT.Abkü = "DIGO" Or lT.Abkü = "DIGOX")
  Case LA_DIGI:
   obLabI = (lT.Abkü = "DIGI" Or lT.Abkü = "DIGIT")
  Case LA_TSH:
   obLabI = (lT.Abkü = "TSH" Or lT.Abkü = "TSHE" Or lT.Abkü = "TSH-L_K" Or lT.Abkü = "TSBL" Or lT.Abkü = "TSBF" Or lT.Abkü = "TSH2" Or lT.Abkü = "TSHTRH" Or lT.Abkü = "XTSH")
  Case LA_fT4:
   obLabI = (lT.Abkü Like "FT4*" Or lT.Abkü = "XFT4")
  Case LA_ft3:
   obLabI = (lT.Abkü Like "FT3*" Or lT.Abkü = "XFT3")
  Case LA_Hb:
   obLabI = (lT.Abkü = "HB" Or lT.Abkü = "XHB")
  Case LA_FOL:
   obLabI = (lT.Abkü = "FOL" Or lT.Abkü = "FOL_S" Or lT.Abkü = "FOLS" Or lT.Abkü = "FOLS_K" Or lT.Abkü = "FOLS01" Or lT.Abkü = "FOLSN")
  Case LA_LEUK:
   obLabI = (lT.Abkü = "LEU" Or lT.Abkü = "LEUK" Or lT.Abkü = "LEUKO" Or lT.Abkü = "LEUK_H" Or lT.Abkü = "LEUK_T" Or lT.Abkü = "XLEUKO")
  Case LA_CRP:
   obLabI = (lT.Abkü = "CRP" Or lT.Abkü = "CRPF" Or lT.Abkü = "CRPHS" Or lT.Abkü = "CRPQ" Or lT.Abkü = "CRPQ01" Or lT.Abkü = "CRPST" Or lT.Abkü = "XCRP")
  Case LA_CK:
   obLabI = (lT.Abkü = "CK" Or lT.Abkü = "CKE" Or lT.Abkü = "CKNAC")
  Case LA_B12:
'   obLabI = (lt.Abkü = "B12" OR lt.Abkü = "B12N" OR lt.Abkü = "B12akt" OR lt.Abkü = "VB12" OR lt.Abkü = "VITB12")
   obLabI = (lT.Abkü = "B12" Or lT.Abkü = "VITB12" Or lT.Abkü = "B12N" Or lT.Abkü = "XVITB12")
  Case LA_THR:
   obLabI = (lT.Abkü = "THRB" Or lT.Abkü = "THRC" Or lT.Abkü = "THRM" Or lT.Abkü = "THRO" Or lT.Abkü = "THRO_E" Or lT.Abkü = "THRO-C" Or lT.Abkü = "THRO-THX" Or lT.Abkü = "THROC" Or lT.Abkü = "THROM" Or lT.Abkü = "THROMB" Or lT.Abkü = "TROM")
  Case LA_GOT:
   obLabI = (lT.Abkü = "GOT" Or lT.Abkü = "GOTR02" Or lT.Abkü = "GOTR03" Or lT.Abkü = "OT" Or lT.Abkü = "XGOT")
  Case LA_AlbS:
   obLabI = (lT.Abkü = "ALBS" Or lT.Abkü = "ALBUS" Or (lT.Abkü = "ALBU" And lT.Einheit = "g/l"))
 End Select
End Function ' obLabI%(LA AS LabArt, lt AS labtyp) ' ob Labor interessant
'' nur noch in Niereninsuffizienzpauschalendiabetiker_Click
'Function LabEPatS(la AS LabArt, Pat_id&) AS CString ', Optional minZp$) AS CString
'#If bis0418 THEN
' Const s1$ = "SELECT DATE(zeitpunkt) zp, IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert) Wert FROM `laborneu` ln LEFT JOIN laborkommentar lk ON ln.kommentarvw = lk.kommentarvw WHERE "
' Const s2$ = " AND pat_id = "
' Const s3$ = " UNION SELECT DATE(u.eingang) zp, IF(ISNULL(w.wert),IF(ISNULL(w.kommentar),'',w.kommentar),w.wert) Wert FROM `" & vorsil & "us` u LEFT JOIN " & vorsil & "wert w ON u.refnr = w.refnr WHERE "
'#END IF
' Const s4$ = " GROUP BY zp ORDER BY zp DESC"
' Dim Abkü$
' SELECT CASE la
'  Case HbA1c: Abkü = "`abkü` IN ('HBA1','HBA1C','HbA1c Eigenlabor','HBA1C_','HBAC') AND CAST(CONCAT('0',REPLACE(`wert`,',','.')) AS decimal) < 22 " ' 30.8.14: wegen falsch eingetragener Fremdlabore gestrichen: AND `einheit` = '%'"
'            ' identisch, wg.Kompatibilität nicht genommen: "RLIKE 'hba[c1]' AND einheit = '%'"
'  Case HbA1cmm: Abkü = "`abkü` IN ('HBA1','HBA1C','HbA1c Eigenlabor','HBA1C_','HBAC') AND CAST(`wert` AS decimal) >= 22 " ' 30.8.14: wegen falsch eingetragener Fremdlabore gestrichen: AND `einheit` = '%'"
'  Case Krea: Abkü = "abkü IN ('CREAT','KRE02','KREA','KREA02','KRES') AND einheit = 'mg/dl'"
'  Case AlbCre: Abkü = "((abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND einheit LIKE 'mg/g %') OR (abkü IN ('ALBU','ALBUMU') AND (einheit = 'mg/l' OR einheit = '')))"
'  Case Kali: Abkü = "abkü IN ('K','KALI')"
'  Case Chol: Abkü = "abkü = 'CHOL'"
'  Case Hdl:  Abkü = "abkü IN ('HDL','HDLC')"
'  Case LDL:  Abkü = "abkü LIKE 'LDL%' AND einheit = 'mg/dl'" ' IN ('LDL','LDLB','LDLH','LDLH01','LDLLG','LDLS')
'  Case eGFR: Abkü = "(abkü RLIKE '^gfr|^gfck|^creacl')"
' END SELECT
' SET LabEPatS = New CString
' 'If minZp = "" THEN
'#If bis0418 THEN
' LabEPatS.AppVar Array(s1, Abkü, s2, Pat_id, s3, Abkü, s2, Pat_id, s4)
'#END IF
' 'Else
''    LabEPatS.AppVar Array(s1, Abkü, s2, Pat_id, " AND zeitpunkt >= ", minZp, s3, Abkü, s2, Pat_id, " AND u.eingang >= ", minZp, s4)
'' END IF
'Const Werte$ = "pat_id, Zeitpunkt zp, abkü, abkü, Langtext,CONVERT(REPLACE(CONCAT('0',Wert),',','.'),DECIMAL(10,1)) Wert,Einheit,Einheit_ur,Kommentar,NB,NB_ur,uNg,uNg_ur,oNg,oNg_ur,Labor,Pfad "
' LabEPatS.AppVar Array("SELECT * FROM (SELECT ", Werte, " FROM `labor2a` WHERE pat_id = ", Pat_id, " AND ", Abkü, " UNION SELECT ", Werte, " FROM `labor1a` WHERE pat_id = ", Pat_id, " AND ", Abkü, " ) i", s4)
''  sql = "SELECT pat_id, zeitpunkt zp, wert FROM `laborneu` ln WHERE abkü RLIKE 'hba[c1]' AND einheit = '%' AND pat_id = " & Pat_id & " UNION SELECT u.pat_id, u.eingang zp, w.wert FROM `" & vorsil & "us` u LEFT JOIN " & vorsil & "wert w ON u.refnr = w.refnr WHERE abkü RLIKE 'hba[c1]' AND einheit = '%' AND pat_id = " & Pat_id & " ORDER BY pat_id,zp DESC"
'End FUNCTION ' LabEPatS

' FUNCTION latinis$(artSpez$)
'  latinis = replace$(artSpez, ",'", ",_latin1'")
'  IF Left$(artSpez, 1) = "'" THEN latinis = "_latin1" & latinis
' END FUNCTION ' latinis$(artSpez$)

Public Sub rrpruef()
 Dim rs As New ADODB.Recordset
 Dim ad As DMPClass
 Dim raAna As New ADODB.Recordset
 Lese.ProgStart
 myFrag rs, "SELECT pat_id, gesname(pat_id) Name,(SELECT MAX(zeitpunkt)from rr WHERE pat_id=n.pat_id) rr1 FROM namen n ORDER BY pat_id DESC" ' WHERE pat_id <60050
 While Not rs.EOF
'  Debug.Print rs!Pat_id
'  Call DMPString$(rs!Pat_id, ad, , , IIf(ISNULL(rs!RR1), Now(), rs!RR1), 0)
  Set raAna = Nothing
  If obAdo Then
   myFrag raAna, "SELECT patAlter(ana.pat_id) PAlter, ana.* FROM `anamnesebogen` ana WHERE pat_id = " & rs!Pat_ID, adOpenKeyset, DBCn, adLockReadOnly
  Else
'   Call raAna.Open("SELECT patAlter(ana.pat_id) PAlter, ana.* FROM `anamnesebogen` ana WHERE pat_id = " & rs!Pat_id, DBCn, adOpenDynamic, adLockReadOnly)
   myFrag raAna, "SELECT patAlter(ana.pat_id) PAlter, ana.* FROM `anamnesebogen` ana WHERE pat_id = " & rs!Pat_ID
  End If
  ad.RRsyst = 0
  ad.RRdiast = 0
  ad.PrRR = GetPrRR(rs!Pat_ID, raAna, ad.RRsyst, ad.RRdiast, obdiastkorr:=True)
  If ad.RRsyst > 300 Or ad.RRsyst < 50 Or ad.RRdiast > 180 Or ad.RRdiast < 30 Then
   If Not IsNull(rs!RR1) Then
    MsgBox "falscher Blutdruck: " & ad.RRsyst & " " & ad.RRdiast
    Stop
   End If
  End If
  ad.RRsyst = 0
  ad.RRdiast = 0
  rs.MoveNext
 Wend
' Debug.Print "Fertig mit rrpruef"
End Sub ' rrpruef

' in DMPAusgeb0, do_DMPAusgebStandAlone, dodoPLZ, DMPAusgeb0, doCallDMP
Function DMPString$(pid&, üdt As DMPClass, Optional ohneVorDMP%, Optional mitBezeich%, Optional DokuDat As Date, Optional mitStr% = True)
 Dim ErgebDatei$, AspZul$, AspBef$, UzuPm$
 Dim Vgst As Date
 
' Dim aktDC.obglib AS AntidiabMedType, aktDC.obmetf AS Boolean, aktDC.obgluci AS Boolean, aktDC.obshglin AS Boolean, aktDC.obglit AS Boolean, aktDC.obsonstAD AS boolean, _
     aktDC.obins AS Boolean, aktDC.obanal AS Boolean, aktDC.obhmg AS Boolean, aktDC.obantihyp AS Boolean, aktDC.obthro AS Boolean
'Dim rEi AS DAO.Recordset, rLaU AS DAO.Recordset, rKH AS DAO.Recordset
Dim raEi As New ADODB.Recordset, raKH As New ADODB.Recordset
'Dim raLau As New ADODB.Recordset, raLaum As New ADODB.Recordset
Dim Labs As labtyp, LabSm As labtyp
Dim rUlc As New ADODB.Recordset
Dim rfuss As New ADODB.Recordset
Dim rfal As New ADODB.Recordset
Dim lapp As New ADODB.Recordset
'Dim ralauSql$
Dim aktDC As DMPClass
'Dim rdo AS DAO.Recordset
'Dim rado As New ADODB.Recordset
Dim raAna As New ADODB.Recordset
Dim raDT As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rsAnam As New ADODB.Recordset
Dim DMSchulz%, HSchulz%
Dim Tbl
'Dim lsql$
'Dim RRsyst%, RRdiast%
Dim fiabfr$
'Dim obVorb% ' ob vorbefund da
' Beim ersten Aufruf von diI muß Pid enthalten sein
' Dim aktdc.ab315%, aktdc.ab317%, aktdc.ab1023%

If Not DokuDat Then DokuDat = MINvb(Now(), fctQEnd(ZQuart(Now() - Verspätung)))
aktDC.x_DokuDat = DokuDat
ErgebDatei$ = aVerz & "DMP.txt"
aktDC.ab315 = (DokuDat > #6/30/2015#)
aktDC.ab317 = (DokuDat > #6/30/2017#)
aktDC.ab1023 = (DokuDat > #9/30/2023#)
aktDC.KtrAbrB = "00"

nochmal:
If diI("Z34 Z33", pid) Then aktDC.obSchw = True
If diI("E10") Then aktDC.dtyp = "1" Else If diI("E11") Then aktDC.dtyp = "2"
If Not ohneVorDMP Then
'fiabfr = "SELECT Pat_ID, FID, form_abk,Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, form_abk FROM (((`forminhfeld` LEFT JOIN `forminhkopf` ON `forminhfeld`.foid=`forminhkopf`.foid) LEFT JOIN `formulare` ON `formulare`.formid=`forminhkopf`.form_id) LEFT JOIN `forminhaltfeld` ON `forminhfeld`.feldvw=`forminhaltfeld`.feldvw) LEFT JOIN `forminhaltfeldinh` ON `forminhfeld`.feldinhvw=`forminhaltfeldinh`.feldinhvw WHERE form_abk LIKE ""dmpdtyp%"" AND feldinh = ""X"" AND pat_id = " & Pid & " ORDER BY zeitpunkt"
fiabfr = "SELECT * FROM (SELECT Pat_ID, FID, form_abk, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh FROM `forminhkopf` LEFT JOIN `formulare` ON `formulare`.formid = `forminhkopf`.form_id LEFT JOIN `forminhfeld` ON `forminhfeld`.foid = `forminhkopf`.foid LEFT JOIN `forminhaltfeld` ON `forminhfeld`.feldvw=`forminhaltfeld`.feldvw LEFT JOIN `forminhaltfeldinh` ON `forminhfeld`.feldinhvw=`forminhaltfeldinh`.feldinhvw AND feldinh='X' WHERE pat_id = " & pid & " AND form_abk LIKE 'dmpdtyp%') AS i WHERE feldinh<>null ORDER BY zeitpunkt"
'raAna.Open fiabfr, DBCn, adOpenDynamic, adLockReadOnly
myFrag raAna, fiabfr, adOpenStatic
Do While Not raAna.EOF ' kommt alle bei uns in Turbomed nicht vor 2.3.25
 aktDC.obVorb = True
 Vgst = raAna!Vorgestellt
 Select Case raAna!Feld
  Case "Folgeerkrankung"
   aktDC.FE(raAna!FeldNr) = True
  Case "SchulEmpfohlen" ' 0 = Diabetes, 1 = Hypertonie, 2 = keine
   aktDC.SE(raAna!FeldNr) = raAna!Zeitpunkt
  Case "SchulWahrgenommen" ' Schulungen bereits vor Einschreibung wahrgenommen: 0 = Diabetes, 1 = Hypertonie, 2 = keine
   aktDC.SW(raAna!FeldNr) = raAna!Zeitpunkt
  Case "DiabSchulung" ' empfohlene Schulungen wahrgenommen, Diabetes: 0 = ja, 1 = nein, 2 = war aktuell nicht möglich
   aktDC.Ds(raAna!FeldNr) = raAna!Zeitpunkt
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
 myFrag rsAnam, "SELECT gesnameg(pat_id) gn, n.* FROM `namen` n WHERE pat_id = " & pid, adOpenStatic, DBCn, adLockReadOnly
Else
 myFrag rsAnam, "SELECT gesnameg(pat_id) gn, n.* FROM `namen` n WHERE pat_id = " & pid, adOpenStatic
End If
If Not rsAnam.BOF Then
aktDC.x_GesName = rsAnam!gn ' GesNamFn(rsAnam) & ", *" & Format$(aktDC.GebDat, "dd.mm.yyyy")
aktDC.Nachname = rsAnam!Nachname
aktDC.Vorname = rsAnam!Vorname
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
aktDC.geschlecht = rsAnam!geschlecht
aktDC.Anschrzus = rsAnam!Anschrzus
aktDC.PrivatTel = rsAnam!PrivatTel
aktDC.KVKStatus = rsAnam!KVKStatus
aktDC.strasse = rsAnam!strasse
aktDC.Hausnr = rsAnam!Hausnr
aktDC.Versichertennummer = Trim$(rsAnam!Versichertennummer)
End If
If obAdo Then
' raAna.Open "SELECT patAlter(ana.pat_id) PAlter, ana.* FROM `anamnesebogen` ana WHERE pat_id = " & pid, DBCn, adOpenKeyset, adLockReadOnly
 myFrag raAna, "SELECT patAlter(ana.pat_id) PAlter, ana.*, COALESCE(MAX(Ulcera),'')='' AND COALESCE(MAX(`Puls Atp`),'')='' AND COALESCE(MAX(`Puls Adp`),'')='' AND COALESCE(MAX(`Vibration IK`),'')='' AND COALESCE(MAX(`Vibration Großzehe`),'')='' AND COALESCE(MAX(`Kalt-Warm`),'')='' anaNdok FROM `anamnesebogen` ana WHERE pat_id = " & pid, adOpenStatic, DBCn, adLockReadOnly
' rsAnam.Open "SELECT * FROM `namen` WHERE pat_id = " & pid, DBCn, adOpenKeyset, adLockReadOnly
Else
' SET rEi = TabÖff("eintraege", "Auswahl")
''Set rlau = TabÖff("Labor", "WertSuch")
' SET rDT = TabÖff("Diagnosen")
' SET rKH = TabÖff("KHEinweis", "Auswahl")
' SET rsAnam = TabÖff("Anamnesebogen", "pat_id")
' Call raEi.Open("SELECT * FROM `eintraege` WHERE pat_id = " & PID, DBCn, adOpenDynamic, adLockReadOnly)
 'Call raDT.Open("SELECT * FROM diagnosen WHERE pat_id = " & PID, DBCn, adOpenDynamic, adLockReadOnly)
 'Call raKH.Open("SELECT * FROM kheinweis WHERE pat_id = " & PID, DBCn, adOpenDynamic, adLockReadOnly)
' Call raAna.Open("SELECT -dialyse AS j_dialyse, a.* FROM `anamnesebogen` a WHERE pat_id = " & pid, DBCn, adOpenDynamic, adLockReadOnly)
'  Call raAna.Open("SELECT patAlter(ana.pat_id) PAlter, ana.* FROM `anamnesebogen` ana WHERE pat_id = " & pid, DBCn, adOpenDynamic, adLockReadOnly)
 myFrag raAna, "SELECT patAlter(ana.pat_id) PAlter, ana.*, COALESCE(MAX(Ulcera),'')='' AND COALESCE(MAX(`Puls Atp`),'')='' AND COALESCE(MAX(`Puls Adp`),'')='' AND COALESCE(MAX(`Vibration IK`),'')='' AND COALESCE(MAX(`Vibration Großzehe`),'')='' AND COALESCE(MAX(`Kalt-Warm`),'')='' anaNdok,`Diabetes seit`,COALESCE(`Jahr letzte Diabetesschulung`,'') RLIKE '^$|^[0nu?-]$|^[-?][^0-9]|nicht|keine|nein|nie|fraglich|^noch|^nur|^kurz|^unklar|^unbekannt' Schulfehlt FROM `anamnesebogen` ana WHERE pat_id = " & pid, adOpenDynamic
' Call rsAnam.Open("SELECT * FROM `namen` WHERE pat_id = " & pid, DBCn, adOpenDynamic, adLockReadOnly)
' DoCmd.Save acForm, Anmnb geht leider nicht
'If PID = 0 THEN
' Dim tonRunde%
' For tonRunde = 1 To 10
'  Call Sound(WinDir + "\media\Windows XP-Standard.wav")
' Next tonRunde
' MsgBox "Achtung: DMP-Infos mit PID=0!"
' ON Error Resume Next
' PID = Forms!Anamnesebogen!Pat_id
' ON Error GoTo fehler
' IF PID = 0 THEN
'  For Each tbl In aktDCb.TableDefs
'   DoCmd.Close acTable, tbl.Name, acSaveYes
'  Next
'  DoCmd.OpenForm aktDCb.Containers(2).Documents(0).Name
'  DoCmd.Maximize
'  PID = Forms!Anamnesebogen!Pat_id
' END IF
'END IF
End If
With raAna
 Dim Herzinfarkt$, Ulcera$, PAlter#, anaNdok%, Diagnosen$, Dialyse&, Größe#, Amputation$, Gewicht$, DiabTyp$
 anaNdok = !anaNdok
 Herzinfarkt = IIf(IsNull(!Herzinfarkt), "", !Herzinfarkt)
 Ulcera = IIf(IsNull(!Ulcera), "", !Ulcera)
 PAlter = IIf(IsNull(!PAlter), 0, !PAlter)
 Diagnosen = IIf(IsNull(!Diagnosen), 0, !Diagnosen)
 Dialyse = IIf(IsNull(!Dialyse), 0, !Dialyse)
 Größe = IIf(IsNull(!Größe), 0, !Größe)
 Amputation = IIf(IsNull(!Amputation), "", !Amputation)
 Gewicht = IIf(IsNull(!Gewicht), 0, !Gewicht)
 DiabTyp = IIf(IsNull(!Diabetestyp), "", !Diabetestyp)
 
AspZul = tfeld(![augensp zuletzt])
AspBef = tfeld(![augensp befund])
UzuPm = tfeld(![Unterzucker pM])

aktDC.Pat_ID = pid
aktDC.SW(0) = Not !Schulfehlt
'If IsNull(!Nachname) Then ' so am 2.9.08 ereignet
' aktDC.Nachname = rsAnam!Nachname
' aktDC.Vorname = rsAnam!Vorname
'Else
' aktDC.Nachname = !Nachname
' aktDC.Vorname = !Vorname
'End If
Pat_ID = pid ' Pat_id = modulübergreifend

If aktDC.Lkz = "" Then aktDC.Lkz = "D"
aktDC.Postleitzahl = "D " & aktDC.plz & " " & aktDC.ort ' 1.1.15 Länderkennzeichen

#If False Then ' 29.6.15, Woltmann
 myFrag rform, "SELECT " & IIf(Not LVobMySQL, "top 1", "") & " feldinh FROM `formular` WHERE pat_id = " & Pat_ID & " AND Feld = 'KVKGueltig' AND zeitpunkt <= " & DatFor_k(MINvb(Now(), fctQEnd(ZQuart(Now - Verspätung)))) & " AND feldinh LIKE '%/%'" & " ORDER BY zeitpunkt DESC" & IIf(LVobMySQL, " LIMIT 1", "")
 If Not rform.BOF Then
  Postleitzahl = left$(Postleitzahl & Space$(25), 25) & rform!FeldInh
 Else
  Postleitzahl = left$(Postleitzahl & Space$(30), 30)
 End If
#End If

 
 myFrag rfal, "SELECT " & IIf(Not LVobMySQL, "top 1", "") & " f.*,LPAD(bPerG,2,'0')bPerG,COALESCE(IKs.IKPrae,'10')IKPrae,COALESCE(kv,'71')KVBereich FROM `faelle` f LEFT JOIN IKs USING(IK) LEFT JOIN plz USING(plz) WHERE pat_id = " & pid & " AND bhfb <= " & DatFor_k(DokuDat) & " ORDER BY bhfb DESC, schgr" & IIf(LVobMySQL, " LIMIT 1", "")
 If Not rfal.BOF Then
  aktDC.VKNr = rfal!VKNr
  aktDC.KVBereich = rfal!KVBereich
  aktDC.IKPrae = rfal!IKPrae
  aktDC.IK = rfal!IK
  aktDC.IKName = rfal!KKasse_2
  aktDC.VschBeg = rfal!VschBeg
  aktDC.VschEnd = rfal!VschEnd
  aktDC.ausgst = rfal!ausgst
  aktDC.lVorl = rfal!lVorl
  aktDC.KtrAbrB = IIf(rfal!KtrAbrB = "", "00", rfal!KtrAbrB)
  aktDC.bPerG = rfal!bPerG
  aktDC.DMPKnZ = rfal!DMPKnZ
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
  aktDC.dmpHbA1cZiel = REPLACE$(rfal!dmpHbA1cZiel, ".", ",")
  aktDC.dmpUewFuss = LCase$(rfal!dmpUewFuss)
  aktDC.dmpEinwDM = LCase$(rfal!dmpEinwDM)
 End If ' Not rfal.BOF Then
 If Not IsNumeric(aktDC.dmpHbA1cZiel) Then
  Dim rHb As New ADODB.Recordset
  myFrag rHb, "SELECT REPLACE(dmphba1cziel,'.',',') FROM faelle WHERE pat_id = " & pid & " AND dmphba1cziel REGEXP '^-?[+]?[.,0-9]+$' ORDER BY bhfb DESC"
  If Not rHb.BOF Then
   aktDC.dmpHbA1cZiel = rHb.Fields(0)
  End If
  Set rHb = Nothing
 End If ' Not IsNumeric(aktDC.dmpHbA1cZiel) Then
 aktDC.dmphalbj = "n" ' LCase$(rfal!dmphalbj)
 Dim rDMPh As New ADODB.Recordset
  myFrag rDMPh, "SELECT MAX(dmphalbj) FROM faelle f0 WHERE pat_id = " & pid & " AND bhfb = (SELECT MAX(bhfb) FROM faelle f WHERE f.pat_id = " & pid & " AND dmphalbj<>'')"
  If Not rDMPh.BOF Then If Not IsNull(rDMPh.Fields(0)) Then aktDC.dmphalbj = rDMPh.Fields(0)
 Set rDMPh = Nothing
'Dim sqllabor$
'sqllabor$ = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, Abkü, Langtext, Wert, Einheit, Kommentar, AbsPos, AktZeit FROM `laborlangtext` INNER JOIN (laborkommentar INNER JOIN `laborneu` ON laborkommentar.KommentarVW = `laborneu`.KommentarVW) ON `laborlangtext`.LangtextVW = `laborneu`.LangtextVW WHERE pat_id = " & CStr(pid)
''Const sqllabor$ = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, Abkü, Langtext, Wert, Einheit, Kommentar, AbsPos, AktZeit FROM `laborlangtext` INNER JOIN (`laborneu`) ON `laborlangtext`.LangtextVW = `laborneu`.LangtextVW"
'lsql = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" AS NB FROM (" & sqllabor & ") AS labor WHERE wert <> """" AND NOT ISNULL(wert)"
'lsql = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" AS NB FROM (" & sqllabor & ") AS labor WHERE wert <> """" AND NOT ISNULL(wert) UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar,Normbereich AS NB " + _
 "FROM `" & vorsil & "us` LEFT JOIN " & vorsil & "wert ON " & vorsil & "us.RefNr=" & vorsil & "wert.RefNr " + _
 "WHERE pat_id = " & aktDC.Pat_id & " AND wert <> """" AND NOT ISNULL(wert) AND NOT EXISTS (SELECT * FROM `laborneu` WHERE pat_id = " & aktDC.Pat_id & " AND abkü = " & vorsil & "wert.Abkü AND wert = " & vorsil & "wert.wert AND zeitpunkt > " & vorsil & "us.Eingang-3 AND zeitpunkt < " & vorsil & "us.Eingang+6)"
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
aktDC.x_DokName = aktDC.x_GesName & " (Pat'nr " & pid & ")" & IIf(mitBezeich, ", DMP-Informationen für " & Format(DokuDat, "d.m.yy"), vNS)
aktDC.x_DoklangName = REPLACE$(aktDC.x_DokName, "*", "geb. ")
 If mitStr Then DmPStrS = IIf(mitBezeich, vNS, "DMP-Informationen vom " & Format$(Date, "DD.MM.YY") & " zu " & vbTab & vbTab & vbTab & vbTab & vbCrLf) & aktDC.x_DokName + ":"
 
 aktDC.x_DmTyp = "D.m. Typ " & dtyp(DiabTyp) & IIf(aktDC.dtyp <> DiabTyp, " (laut Anamnesebogen; " & aktDC.dtyp & " laut ICD)", vNS) & " seit:"
 If Not IsNull(Vgst) Then
  aktDC.daseit = zuJahr(DSeit(raAna, True), Vgst)
  If mitStr Then TabPr aktDC.x_DmTyp, aktDC.daseit
 End If ' IsNull(vgst) Then
 
 #If vorca2008 Then
  If obPosi(![Unterzucker pM]) Or obPosi(![keto pa]) Or obPosi(![Fremde Hilfe pa]) Or obPosi(![Bewußtlos pa]) Then
   aktDC.dspsy = True
   If mitStr Then TabPr "Diagnosespez. Symptome:", "  ja"
  End If
  If diI("Z34 Z33") Then
   If mitStr Then TabPr "Schwangerschaft laut ICD-Diagnosen:", "  ja"
   aktDC.obSchw = True
  End If
 #End If ' vorca2008
 
 Dim VorDat As Date '(nur Attrappe)
 Call TherAuskunft(CStr(pid), 0, aktDC.insz, VorDat, aktDC.obIns, aktDC.obAnal, aktDC.obGlib, aktDC.obmetf, aktDC.obGlucI, aktDC.obSHGlin, aktDC.obGlit, aktDC.obDpp4, aktDC.obGlp1, aktDC.obSglt2, aktDC.obSonstAD, aktDC.obHMG, aktDC.obAntihyp, aktDC.obACEH, aktDC.obBetabl, aktDC.obThro, aktDC.obOAK, , , , aktDC.obDiur, aktDC.obAT1)
 'if mitstr THEN TabPr "Diagnosespez. Medik.: ", IIf(!`DiabetesMedikament 1` <> "" AND NOT ISNULL(!`DiabetesMedikament 1`), "  ja", "nein")
 If aktDC.dtyp = "1" And Not aktDC.obIns And Not aktDC.obAnal Then ' Medikamentenplan fehlt oder keine Dosierung angegeben
  Dim rez As New ADODB.Recordset, rezSQL$
  rezSQL = "SELECT * FROM (SELECT *,LEFT(medikament,instr(medikament,"" "")-1) AS medanf FROM `rezepteintraege` ) AS rez LEFT JOIN `medarten` ON rez.medanf = `medarten`.medikament WHERE rez.pat_id = 150 AND (ins OR anal) ORDER BY zeitpunkt DESC;"
'  IF Not lies.obMySQL THEN rezSQL = replace$(rezSQL, "collate latin1_german2_ci ", vNS)
  myFrag rez, rezSQL
  If Not rez.BOF Then
   aktDC.obIns = rez!InS
   aktDC.obAnal = rez!anal
  End If
  rez.Close
 End If
 
 #If vorca2008 Then
  aktDC.dspmed = (aktDC.obIns Or aktDC.obAnal Or aktDC.obGlib = adja Or aktDC.obmetf = adja Or aktDC.obGlucI = adja Or aktDC.obSHGlin = adja Or aktDC.obGlit = adja Or aktDC.obDpp4 = adja Or aktDC.obGlp1 = adja Or aktDC.obSglt2 = adja Or aktDC.obSonstAD = adja)
' aktDC.dspmed = (aktDC.VorM(0) OR aktDC.VorM(1) OR aktDC.VorM(2) OR aktDC.VorM(3) OR aktDC.VorM(4) OR aktDC.VorM(5) OR aktDC.VorM(6))
  If mitStr Then TabPr "Diagnosespez. Medik.:", IIf(aktDC.dspmed, "  ja", "nein")
 #End If ' vorca2008
 
' 2.Reiter
 Dim spät As New CString, ereig As New CString, bbk As New CString
 'If instrb(!Diagnosen, "yperton") <> 0 OR instrb(!Diagnosen, "ochdru") <> 0 THEN bbk.append " Hypertonus,"
 'If diT("yperton") OR diT("ochdru") THEN bbk.append " Hypertonus,"
 If diI("I10") Then
  aktDC.FEn(1) = True
  bbk.AppVar Array(" Art.Hypertonie", IIf(aktDC.FE(1) And Not aktDC.ab315, "( bek)", ""), ",")
 End If
 If diI("I50") Then ' Herzinsuffizienz
  aktDC.FEn(17) = True
  bbk.AppVar Array(" Chron.Herzinsuff.", IIf(aktDC.FE(17) And Not aktDC.ab315, "( bek)", ""))
 End If
 If diT("ettstoffw yperchol") Or diI("E78") Then
  aktDC.FEn(2) = True
  bbk.AppVar Array(" Fettstoffw'strg.", IIf(aktDC.FE(2) And Not aktDC.ab315, "( bek)", ""), ",")
 End If
 If diI("J45") Then ' Asthma
  aktDC.FEn(16) = True
  bbk.AppVar Array(" Asthma bronch.", IIf(aktDC.FE(16) And Not aktDC.ab315, "( bek)", ""), "")
 End If
 If diI("J42 J44") Then ' COPD
  aktDC.FEn(15) = True
  bbk.AppVar Array(" COPD", IIf(aktDC.FE(15) And Not aktDC.ab315, "( bek)", ""), "")
 End If
 If diI("I7") Then ' pAVK usw
  aktDC.FEn(6) = True
  bbk.AppVar Array(" pAVK", IIf(aktDC.FE(6) And Not aktDC.ab315, "( bek)", ""), ",")
 End If
 If diI("I20 I21 I22 I23 I24 I25") Then ' I20 A.p., I21 akuter MI, I22 rez.MI, I23 Kompl.n.MI, I24 sonst.ischäm.Hkt, I25 Atherosk.Hkt, alter MI
  aktDC.FEn(3) = True
  bbk.AppVar Array(" Koronare Herzerkrankung", IIf(aktDC.FE(3) And Not aktDC.ab315, "( bek)", ""), ",")
 End If
  
 If Not aktDC.ab315 Then If LenB(bbk) <> 0 Then bbk.AppVar Array(vbCrLf, Space$(BreiteSp1 - 1))
 Dim ICD$, diagdat As Date
 If obPosi(Herzinfarkt) Or diI("I21 I22", , , True, , aktDC.FEi(4), aktDC.FEd(4)) Then
  aktDC.FEn(4) = True
  If Not aktDC.ab315 Then bbk.AppVar Array(" Herzinfarkt", IIf(aktDC.FE(4) And Not aktDC.ab315, "( bek)", ""), ",")
 End If
 If diT("poplex chlaganf erebral", , , True, , aktDC.FEi(5), aktDC.FEd(5)) Or _
    diI("I63 I64", , , True, , aktDC.FEi(5), aktDC.FEd(5)) Then
  aktDC.FEn(5) = True
  If Not aktDC.ab315 Then bbk.AppVar Array(" Schlaganfall", IIf(aktDC.FE(5) And Not aktDC.ab315, "( bek)", ""), ",")
 End If
 If diI("N08.3 N18 N19") And Not obkNeph Then
  aktDC.FEn(7) = True
  If aktDC.ab315 Then
   spät = "Nephropathie"
  Else
   bbk.AppVar Array(" Nephropathie", IIf(aktDC.FE(7) And Not aktDC.ab315, "( bek)", ""), ",")
  End If
 End If
  
' IF obPosi(!j_Dialyse) OR diI("Z49.1") THEN
 If obPosi(Dialyse <> 0) Or diI("Z49", , , , , aktDC.FEi(8), aktDC.FEd(8)) Then
  aktDC.FEn(8) = True
  If aktDC.ab315 Then
   spät.AppVar Array(IIf(spät.Length = 0, "Nephropathie:", ":"), " Nierenersatztherapie")
  Else
   bbk.AppVar Array(" Nierenersatztherapie", IIf(aktDC.FE(8) And Not aktDC.ab315, "( bek)", ""), ",")
  End If
 End If
 
  aktDC.obG632 = diI("G63.2")
  If aktDC.obG632 Or diI("G59.0 G99.0") Or diT("europath") Then
   aktDC.FEn(11) = True
   If aktDC.ab315 Then
    spät.AppVar Array(IIf(spät.Length = 0, "", ", "), "Neuropathie")
   End If
  End If
  
  If diT("Retinopath") Or diI("H36") Then
   aktDC.FEn(9) = True
   If aktDC.ab315 Then
    spät.AppVar Array(IIf(spät.Length = 0, "", ", "), "Retinopathie")
   Else
    bbk.AppVar Array(" diab. Retinopathie", IIf(aktDC.FE(9) And Not aktDC.ab315, "( bek)", ""), ",")
   End If
  End If
  If InStrB(LCase$(Diagnosen), "blind") <> 0 Or diI("H54 S05", , , True, , aktDC.FEi(10), aktDC.FEd(10)) Then
   aktDC.FEn(10) = True
   If aktDC.ab315 Then
    spät.AppVar Array(IIf(aktDC.FEn(9), ", ", IIf(spät.Length = 0, "", ", ") & "Retinopathie"), " (Blindheit)")
   Else
    bbk.AppVar Array(" Blindheit", IIf(aktDC.FE(10) And Not aktDC.ab315, "( bek)", ""), ",")
   End If
  End If
      
  If Not aktDC.ab315 Then
   If LenB(bbk) <> 0 Then bbk.AppVar Array(vbCrLf, Space$(BreiteSp1 - 1))
   If aktDC.FEn(11) Then
    bbk.AppVar Array(" diab. Neuropathie", IIf(aktDC.FE(11) And Not aktDC.ab315, "( bek)", ""), ",")
   End If
  End If
  
  If diI("L89 M14.6") Then
   aktDC.FEn(12) = True
   If Not aktDC.ab315 Then bbk.AppVar Array(" Diabetisches Fußsyndrom", IIf(aktDC.FE(12) And Not aktDC.ab315, "( bek)", ""), ",")
  End If
  If obPosi(Amputation) Or diT("mputat", , , True, , aktDC.FEi(13), aktDC.FEd(13)) Then
   aktDC.FEn(13) = True
   If Not aktDC.ab315 Then bbk.AppVar Array(" Amputation", IIf(aktDC.FE(13) And Not aktDC.ab315, "( bek)", ""), ",")
  End If

  If diI("M36.8 M14.2") Or InStrB(Diagnosen, "aculop") <> 0 _
  Or diI("K77.8 K71.7 K71.6 K76.9 K76.2") Or diT("atara") > 0 Or diT("arnwegsinf") > 0 Then
   aktDC.FEn(14) = True
   If Not aktDC.ab315 Then bbk.AppVar Array(" Sonstige", IIf(aktDC.FE(14) And Not aktDC.ab315, "( bek)", ""), "")
  End If
  
 Dim j%
 aktDC.FEn(0) = -1
 For j = 1 To UBound(aktDC.FEn)
  If aktDC.FEn(j) And Not aktDC.FE(j) Then
   aktDC.FEn(0) = 0
   Exit For
  End If
 Next j
 If bbk.Length = 0 Then bbk = " keine"
 If Right$(bbk, 1) = "," Then bbk.Cut (bbk.Length - 1)
 Do While Asc(bbk.Right(1)) = 9 Or Asc(bbk.Right(1)) = 10 Or Asc(bbk.Right(1)) = 11 Or Asc(bbk.Right(1)) = 13
  bbk.Cut (bbk.Length - 1)
  If Len(bbk) = 0 Then Exit Do
 Loop
 aktDC.x_Bbk = LTrim$(bbk)
 If mitStr Then TabPr "•  Bek. Begleit-/Folgeerk.:", aktDC.x_Bbk
' Dim TabakExpr$
' IF ISNULL(!Tabak) OR !Tabak = "" THEN
'  TabakExpr = vns
' Else
'  IF instrb("jJ", trim$(!Tabak)) > 0 THEN
'   TabakExpr = "  ja"
'  ElseIf instrb("-nN0", trim$(!Tabak)) > 0 THEN
'   TabakExpr = "nein"
'  Else
'   TabakExpr = CStr(!Tabak)
'  END IF
' END IF
 aktDC.obRauch = diI("F17.1")
 aktDC.Tabak = (WieTabak(Pat_ID) = aktuell)
 If mitStr Then TabPr "Raucher:", IIf(aktDC.Tabak, "  ja", "nein")
 If Not IsNull(Größe) Then
  aktDC.kGR = Round(REPLACE$(IIf(Größe < 10, Größe * 100, Größe), ".", ","), 0)
  If mitStr Then TabPr "Körpergröße:", CStr(Größe) + IIf(Größe < 10, " m", " cm")
 End If ' Not IsNull(!Größe) Then
 
' raEi.Open "SELECT * FROM `eintraege` WHERE pat_id = " & pid & " AND art = ""Gewicht"" ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
 myFrag raEi, "SELECT * FROM `eintraege` WHERE pat_id = " & pid & " AND art RLIKE '^gew$|^gw$|^gewi' ORDER BY zeitpunkt DESC"
 If Not raEi.EOF Then
  Dim gewi$
  gewi = Trim$(REPLACE$(REPLACE$(LCase$(raEi!Inhalt), "kg", ""), ".", ","))
  If InStrB(gewi, "-") <> 0 Then gewi = left$(gewi, InStr(gewi, "-") - 1)
   Dim pos&
  Do
   If Not IsNumeric(gewi) Then
    pos = InStr(gewi, " ")
    If pos > 0 Then
     If IsNumeric(left$(gewi, pos - 1)) Then
      gewi = left$(gewi, pos - 1)
     Else 'If IsNumeric(mid$(gewi, pos + 1)) THEN
      gewi = Mid$(gewi, pos + 1)
     End If
    Else
     For j = 33 To 255
      If (j < 48 Or j > 57) And j <> 44 Then
       gewi = REPLACE$(gewi, Chr$(j), "")
      End If
     Next j
    End If
   End If
   If InStr(gewi, ",") > 0 Then
    If InStr(InStr(gewi, ",") + 1, gewi, ",") > 0 Then
     gewi = left$(gewi, InStr(InStr(gewi, ",") + 1, gewi, ",") - 1)
    End If
   End If
   gewi = REPLACE$(gewi, ",,", ",")
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
  aktDC.x_Gewi = raEi!Inhalt + IIf(InStrB(raEi!Inhalt, "kg") <> 0, vNS, " kg") & " (" & CStr(DateValue(raEi!Zeitpunkt)) & ")"
 Else
  If IsNull(Gewicht) Then
   aktDC.x_Gewi = "? kg"
  Else
   aktDC.gewi = Round(REPLACE$(Gewicht, ".", ","), 0) ' ist double
   aktDC.x_Gewi = Gewicht + " kg"
  End If ' IsNull(!Gewicht) Then else
 End If ' Not raEi.EOF THEN else
 If mitStr Then TabPr "Körpergewicht:", aktDC.x_Gewi
    
 If aktDC.kGR <> 0 Then
  aktDC.bmi = aktDC.gewi / (aktDC.kGR * 1# * aktDC.kGR)
  If aktDC.bmi > 0 Then
   Do While aktDC.bmi < 10
    aktDC.bmi = aktDC.bmi * 10
   Loop
   Do While aktDC.bmi > 100
    aktDC.bmi = aktDC.bmi * 0.1
   Loop
  End If
 End If ' aktDC.kgr <> 0 THEN
 
 aktDC.PrRR = GetPrRR(pid, raAna, aktDC.RRsyst, aktDC.RRdiast, obdiastkorr:=True)
 If aktDC.RRdiast < 30 Then aktDC.RRdiast = MINvb(70, aktDC.RRsyst) ' 31.10.20
 If aktDC.dtyp = "2" Then
  If aktDC.PrRR <> "" Then
   If mitStr Then TabPr "Blutdruck:", aktDC.PrRR
  End If ' aktDC.PrRR <> "" Then
 End If ' aktDC.dtyp = "2" Then
 
If True Then ' lwZahl
 aktDC.bekHb = 0
' ralau.Seek "=", CStr(PID), "HBA1C"
 Dim DMPCrea$
' SET raLau = Nothing
' raLau.Open "SELECT * FROM (" & lsql & ") AS innen WHERE abkü = ""HBA1C"" ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
' SET raLau = LabEPat(HbA1c, PID)
 alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
 Labs = LabPat(LA_HbA1c, pid)
' Debug.Print Labs.WertSg
' raLaU.MoveFirst
' raLaU.Find "Abkü = ""HBA1C"""
'  aktdc.x_DMPHbA1c = raLau!Wert & "% (" & Format$(raLau!Zp, "dd/mm/yy") & ")" & IIf(Not aktdc.ab315, ", oberer Normwert: 6,2%", "")
 
' SET raLaum = LabEPat(HbA1cmm, PID)
 alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
 LabSm = LabPat(LA_HbA1cmm, pid)

 Dim Quelle% '0= raLau, 1=raLaum, 2= Formular
' IF Not raLau.BOF AND raLaum.BOF THEN
' ElseIf raLau.BOF AND NOT raLaum.BOF THEN
 If Labs.Abkü = "" And LabSm.Abkü <> "" Then
  Quelle = 1
' ElseIf raLau.BOF AND raLaum.BOF THEN
 ElseIf Labs.Abkü = "" And LabSm.Abkü = "" Then
  Quelle = 2
' ElseIf raLau!zp >= raLaum!zp THEN
' ElseIf raLau!zp < raLaum!zp THEN
 ElseIf Labs.Zp <= LabSm.Zp Then
  Quelle = 1
 Else
  Quelle = 0
 End If
 If Quelle = 1 Then
'  aktdc.x_DMPHbA1c = ROUND((raLaum!wert * 0.0915 + 2.15), 1) & " % / " & raLaum!wert & " mmol/mol (" & Format$(raLaum!zp, "dd/mm/yy") & ")"
  aktDC.bekHb = Round((LabSm.WertSg * 0.0915 + 2.15), 1)
  aktDC.x_DMPHbA1c = aktDC.bekHb & " % / " & LabSm.WertSg & " mmol/mol (" & Format$(LabSm.Zp, "dd/mm/yy") & ")"
 ElseIf Quelle = 0 Then
'  IF raLau!wert <> "" THEN
  If Labs.WertSg <> "" Then
   If InStrB(Labs.WertSg, " ") <> 0 Then
    pos = InStr(Labs.WertSg, " ")
    Dim zwStri$
    zwStri = REPLACE$(left$(Labs.WertSg, pos - 1), ".", ",")
    If IsNumeric(zwStri) Then
     aktDC.bekHb = Round(CDbl(zwStri), 1)
    End If
   Else
    Dim rLauWert$
    rLauWert = REPLACE$(Labs.WertSg, "?", vNS)
    pos = InStr(rLauWert, ">") ' Pat_id 744
    If pos > 0 Then rLauWert = Mid$(rLauWert, pos + 1)
    aktDC.bekHb = Round(CDbl(REPLACE$(rLauWert, ".", ",")), 1)
   End If
  End If
'  aktdc.x_DMPHbA1c = raLau!wert & " % / " & ROUND((aktDC.bekHb - 2.15) * 10.929) & " mmol/mol (" & Format$(raLau!zp, "dd/mm/yy") & ")"
'  aktdc.x_DMPHbA1c = LabS.wertsg & " % / " & ROUND((aktDC.bekHb - 2.15) * 10.929) & " mmol/mol (" & Format$(LabS.zp, "dd/mm/yy") & ")"
  Dim LabsWert!
  LabsWert = IIf(Labs.WertSg = vNS, 0, aktDC.bekHb)
  aktDC.x_DMPHbA1c = LabsWert & " % / " & Round((LabsWert - 2.15) * 10.929) & " mmol/mol (" & Format$(Labs.Zp, "dd/mm/yy") & ")"
 ElseIf Quelle = 2 And tfeld(![letztes HbA1c]) <> "" Then
  Dim testHbA1c$
  testHbA1c = REPLACE$(REPLACE$(![letztes HbA1c], ".", ","), "%", "")
  If IsNumeric(testHbA1c) Then
   aktDC.bekHb = Round(REPLACE$(testHbA1c, ".", ","), 1)
  End If
  aktDC.x_DMPHbA1c = ![letztes HbA1c] & " " & IIf(LenB(tfeld(![gemessen am])) = 0, "(auswärtig)", "( " & tfeld(![gemessen am]) & ")")
 End If
 
 
 ' aktdc.x_DMPHbA1c = aktdc.x_DMPHbA1c & " / " & raLau!Wert & " mmol/l (" & Format$(raLau!Zp, "dd/mm/yy") & ")"
 
' SET raLau = Nothing
' SET raLau = LabEPat(Krea, PID)
 alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
 Labs = LabPat(LA_Krea, pid)
' IF Not raLau.EOF THEN
 If Labs.Abkü <> "" Then
'  aktDC.Crea = ROUND(replace$(replace$(IIf(Right$(raLau!Wert, 4) = "Jaffe", Left$(raLau!Wert, Len(raLau!Wert) - 4), raLau!Wert), ".", ","), ";", ","), 1)
  Dim CreaR$
  CreaR = Labs.WertSg
  If Right$(CreaR, 5) = "Jaffe" Then
   CreaR = left$(CreaR, Len(CreaR) - 5)
  End If
  CreaR = REPLACE$(CreaR, ".", ",")
  CreaR = REPLACE$(CreaR, ";", ",")
  CreaR = REPLACE$(CreaR, "(?)", vNS)
  CreaR = REPLACE$(CreaR, "?", vNS)
  CreaR = REPLACE$(CreaR, "(S)", "")
  If CreaR = "" Then CreaR = "0"
  If InStrB(CreaR, " ") <> 0 Then CreaR = left$(CreaR, InStr(CreaR, " ") - 1)
  If CreaR = "o,B," Then CreaR = "1.0"
  
  Dim pospfeil%
  pospfeil = MAXvb(InStr(CreaR, ">"), InStr(CreaR, "<"))
  If pospfeil > 0 Then CreaR = Mid$(CreaR, pospfeil + 1)
  aktDC.Crea = REPLACE$(REPLACE$(CreaR, "^", ""), ",,", ",")
  aktDC.Crea = Round(aktDC.Crea, 1)
  DMPCrea = Labs.WertSg & " mg/dl (" & Format$(Labs.Zp, "dd/mm/yy") & ")" ', oberer Normwert: 1,3 mg/dl"
 End If
 If aktDC.ab315 Then
 ' das folgende umgeändert am 15.4.20
''  SET raLau = Nothing
''  SET raLau = LabEPat(eGFR, PID)
'  alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
'  Labs = LabPat(LA_eGFR, PID)
''  IF Not raLau.EOF THEN
'  IF Labs.Abkü <> "" THEN
'   Dim eGFRs$
'   eGFRs = Labs.WertSg
'   eGFRs = replace$(eGFRs, ".", ",")
'   eGFRs = replace$(eGFRs, ";", ",")
'   eGFRs = replace$(eGFRs, "(?)", ",")
'   eGFRs = replace$(eGFRs, "?", ",")
'   eGFRs = replace$(eGFRs, "(S)", ",")
'   IF eGFRs = "" THEN eGFRs = "0"
'   IF InStrB(eGFRs, " ") <> 0 THEN eGFRs = Left$(eGFRs, InStr(eGFRs, " ") - 1)
'   IF eGFRs = "o,B," OR eGFRs = ">60" THEN eGFRs = "80"
'   IF eGFRs = ">" THEN aktDC.eGFR = 80 ELSE aktDC.eGFR = CLng(eGFRs)
'
'   aktDC.eGFR = ROUND(aktDC.eGFR, 0)
'   IF aktDC.eGFR = 0 THEN
'    IF aktDC.Crea = 0 THEN
'     aktDC.eGFR = 0
'    Else
'     aktDC.eGFR = ROUND((140 - raAna!palter) * aktDC.gewi * IIf(aktDC.gewi < 3, 100, 1) * IIf(aktDC.Geschlecht = "w", 0.85, 1) / aktDC.Crea / 72, 0)
'    END IF
'   END IF
'   aktdc.x_DMPeGFR = IIf(aktDC.eGFR, aktDC.eGFR & " ml/min (" + Format$(Labs.Zp, "dd/mm/yy") + ")", "nicht vorliegend") ', oberer Normwert: 1,3 mg/dl"
'  END IF
  Dim lG As labtyp
  lG = letztGFR(pid, PAlter, IIf(aktDC.geschlecht = "w", -1, 0))
  lG.WertSg = REPLACE$(lG.WertSg, "<", "")
  aktDC.eGFR = IIf(Not IsNumeric(lG.WertSg), 200, lG.WertSg)
  aktDC.x_DMPeGFR = IIf(aktDC.eGFR, aktDC.eGFR & " " & lG.Einheit & " (" + Format$(lG.Zp, "dd/mm/yy") + ")", "nicht vorliegend") ', oberer Normwert: 1,3 mg/dl"
 End If ' aktdc.ab315 THEN
 
' IF aktDC.dtyp = "1" THEN ' 5.10.08: jetzt auch Typ 2
'  SET raLau = Nothing
'  raLau.Open "SELECT * FROM (" & lsql & ") AS innen WHERE abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND (abkü <> 'ALBU' OR wert LIKE '%<%') ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
'  sql = "SELECT zeitpunkt zp, abkü, einheit, wert FROM `laborneu` WHERE abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND pat_id = " & Pat_id & " UNION SELECT u.eingang zp, w.abkü abkü, w.einheit einheit, w.wert wert FROM `" & vorsil & "us` u LEFT JOIN " & vorsil & "wert w ON u.refnr = w.refnr WHERE abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND pat_id = " & Pat_id & " ORDER BY pat_id,zp DESC"
'  SET raLau = LabEPat(AlbCre, PID)
  alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
  Labs = LabPat(LA_AlbCre, pid)
  ' ALBCRE, ALBKRE, ALBQ, ALBUM, ALBUP
' raLaU.Find "Abkü IN (""CREAT"",""KRE02"", ""KREA"", ""KREA02"", ""KRES"")"
'  IF raLau.EOF THEN
  If Labs.Abkü = "" Then
   Dim rEintr As New ADODB.Recordset
   myFrag rEintr, "SELECT * FROM `eintraege` WHERE art = ""urin"" AND inhalt LIKE ""%micral%"" AND Pat_ID = " & aktDC.Pat_ID & " ORDER BY zeitpunkt DESC"
   If Not rEintr.EOF Then
    aktDC.x_DMPUAlb = Mid$(rEintr!Inhalt, InStr(1, rEintr!Inhalt, "micral", vbTextCompare)) & " (" & Format$(rEintr!Zeitpunkt, "dd/mm/yy") & ")"
   End If ' Not rEintr.EOF Then
  Else ' Labs.Abkü = "" Then
   If InStrB(Labs.WertSg, "<") <> 0 Then
    aktDC.mau = unauff
   Else ' InStrB(Labs.WertSg, "<") <> 0 Then
     Dim raLauW$
     raLauW = REPLACE$(REPLACE$(REPLACE$(IIf(Labs.WertSg = vNS, "0", Labs.WertSg), "ca.", ""), ".", ","), " (Urin)", vNS)
     If IsNumeric(raLauW) Then
      If Round(raLauW, 1) > 20 Then
       aktDC.mau = auff
      Else
       aktDC.mau = unauff
      End If
     Else
      aktDC.mau = unauff
     End If
   End If ' InStrB(Labs.WertSg, "<") <> 0 Then else
   Dim werts$
   If IsNumeric(Labs.WertSg) Then werts = CStr(Labs.WertSg) Else werts = Labs.WertSg
   aktDC.x_DMPUAlb = werts + " mg/dl (" + Format$(Labs.Zp, "dd/mm/yy") + "), oberer Normwert: 29 mg/gCrea"
   If IsNumeric(Labs.WertSg) Then
    If Labs.WertSg >= 30 Then
     aktDC.mau = auff
    Else
     aktDC.mau = unauff
    End If
   End If ' IsNumeric(Labs.WertSg) Then
'  END IF
  End If ' Labs.Abkü = "" Then else
  If aktDC.dtyp = "2" Then
   If mitStr Then TabPr IIf(aktDC.ab317, "•  ", "") & "HbA1c:", aktDC.x_DMPHbA1c
   If aktDC.ab315 Then
    If mitStr Then TabPr "eGFR:", aktDC.x_DMPeGFR
   Else
    If mitStr Then TabPr "Serum-Kreatinin:", DMPCrea
   End If
   If mitStr Then TabPr "Urin-Albumin:", aktDC.x_DMPUAlb ' 5.10.08 Kommentar entfernt
  End If ' aktDC.dtyp = "2" Then
 'If obPosi(!Beinbefund) OR obPosi(Ulcera) OR diI("M14") THEN
 End If 'lwzahl
 
 On Error Resume Next
 ReDim rNa(0)
 On Error GoTo fehler
 rNa(0).Pat_ID = pid
 Call usdmAlt(True)
 Dim trp As Boolean, tlp As Boolean, drp As Boolean, dlp As Boolean
 Dim Mfrep%, Mflip%
 Dim KWrep%, KWlip%
 Dim VibrIKrep%, VibrIKlip%
 Dim VibrGZrep%, VibrGZlip%
 Dim summe%
 Dim MFBegr$, KWBegr$, VibBegr$
 aktDC.Puls = PStatNeu(pid, aktDC.x_Pulsstatus)
 If mitStr And Not aktDC.ab317 Then TabPr "Pulsstatus: ", aktDC.x_Pulsstatus
 aktDC.sens = sensib(pid, aktDC.x_SensText)
 If mitStr And Not aktDC.ab317 Then TabPr "Sensibilität: ", aktDC.x_SensText
   
 If aktDC.ab315 Then
  If spät.Length = 0 Then spät = "keine"
  aktDC.x_Spät = spät
  If mitStr Then TabPr "Spätfolgen: ", aktDC.x_Spät
 End If ' mitStr AND aktdc.ab315 THEN


#If vorca2008 Then
 Dim abiind&
 Dim LUSDat$

' aktDC.Puls = unauff
' For abiind = 2 To 5
'  IF Len(AbI(abiind)) > 0 THEN
'   IF Left$(AbI(abiind), 1) = "{" THEN
'    IF Right$(AbI(abiind), 1) = "}" THEN
'     AbI(abiind) = ""
'    END IF
'   END IF
'  END IF
' Next abiind
' IF UBound(AbI) = 0 OR (AbI(3) = "" AND AbI(2) = "" AND AbI(5) = "" AND AbI(4) = "") THEN
'  PulsParse ![Puls Atp], trp, tlp
'  PulsParse ![Puls Adp], drp, dlp
'  Nausg = " (" & tfeld(![Puls Atp]) & "," & tfeld(![Puls Adp]) & ")"
'  LUSDat = vgst
'  IF tfeld(![Puls Atp]) = "" AND tfeld(![Puls Adp]) = "" THEN aktDC.Puls = ndok
' Else
'  PulsParse AbI(3) & "/" & AbI(2), trp, tlp
'  PulsParse AbI(5) & "/" & AbI(4), drp, dlp
'  Nausg = " (" & AbI(3) & "/" & AbI(2) & ", " & AbI(5) & "/" & AbI(4) & ")"
'  LUSDat = Format$(AbIDate, "d.m.yy")
' END IF
 
 For abiind = 10 To 17
  If Len(AbI(abiind)) > 2 Then
   If left$(AbI(abiind), 1) = "{" Then
    If left$(Right$(AbI(abiind), 3), 2) = "}/" Then
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
' IF (trp AND drp) OR (tlp AND dlp) OR (trp AND dlp) OR (tlp AND drp) THEN
'   aktDC.Puls = auff
' END IF
' IF aktDC.Puls = auff THEN
'  IF mitStr THEN TabPr "Pulsstatus:", "auffällig (" & LUSDat & Nausg
' ElseIf aktDC.Puls = ndok THEN
'  IF mitStr THEN TabPr "Pulsstatus:", "nicht dokumentiert (" & LUSDat & Nausg
' Else
'  IF mitStr THEN TabPr "Pulsstatus:", "unauffällig (" & LUSDat & Nausg
' END IF

 summe = Mfrep + Mflip + KWrep + KWlip + VibrIKrep + VibrIKlip + VibrGZrep + VibrGZlip
 Dim Nausg$
 Nausg = "auffällig (" & LUSDat
 If summe > 1 Then ' OR aktDC.obG632 THEN
  If Mfrep > 0 Or Mflip > 0 Then Nausg = Nausg & " Monofil.:" & MFBegr & ","
  If KWrep > 0 Or KWlip > 0 Then Nausg = Nausg & " Kalt/W.:" & KWBegr & ","
  If VibrGZrep > 0 Or VibrGZlip > 0 Then Nausg = Nausg & " Vibr." & VibBegr & ","
  aktDC.sens = auff
 ElseIf aktDC.obG632 Then
  aktDC.sens = pathdok
  If aktDC.obG632 Then Nausg = Nausg & " Diagn. G63.2"
 ElseIf tfeld(!Monofilamenttest) = "" And (tfeld(![Kalt-Warm]) = "" Or tfeld(![Kalt-Warm]) = "/5 | /5" Or tfeld(![Kalt-Warm]) = "~/5 | ~/5") And tfeld(![Vibration IK]) = "" And tfeld(![Vibration Großzehe]) = "" Then
  aktDC.sens = ndok
  Nausg = "nicht dokumentiert"
 Else
  aktDC.sens = unauff
  Nausg = "un" & left$(Nausg, Len(Nausg) - 2) & ")"
 End If
 If mitStr Then TabPr "Sensibilitätsprüfung:", Nausg
#End If ' vorca2008

' inj: 0 = nicht untersucht, 1 = unauffällig, 2 = auffällig
    Dim injf$
    If aktDC.obIns = 0 And aktDC.obAnal = 0 Then
     aktDC.inj = "0"
    Else
     Dim rusdm As New ADODB.Recordset
     myFrag rusdm, "SELECT spritzst FROM usdm WHERE pat_id = " & pid & " AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM usdm WHERE pat_id = " & pid & " AND zeitpunkt < " & Format(DokuDat + 1, "yyyymmdd") & ")"
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
  If aktDC.ab315 Then
   aktDC.x_Inj = IIf(aktDC.inj = 0, "Nicht untersucht", IIf(aktDC.inj = 1, "Unauffällig", "Auffällig"))
   If mitStr Then If aktDC.ab317 Then TabPr "Injektionsstellen:", aktDC.x_Inj
  End If ' aktdc.ab315 Then
  
' 29.8.17
sql = "SELECT IF(MID(di.icd, 5, 1)>=2,1,IF(MID(di.icd,5,1)<=1,0,IF(u.ulcera LIKE '%obfl%',0,IF(u.ulcera LIKE '%tief%',1,IF(e.art IS NULL,2,0))))) ulcus " & _
"FROM aktfvs f " & _
"LEFT JOIN eintraege e ON f.pat_id = e.pat_id AND e.art='ulcus' AND e.zeitpunkt BETWEEN qanf() AND qend() " & _
"LEFT JOIN fuss u ON f.pat_id = u.pat_id AND u.ulcera IN ('obfl','tief') AND u.zeitpunkt BETWEEN qanf() AND qend() " & _
"LEFT JOIN diagview di ON f.pat_id = di.pat_id AND  di.gICD RLIKE '^L89\.[1-5]' AND obdauer = 0 AND DATE(di.diagdatum) BETWEEN qanf() AND qend() " & _
"WHERE (NOT ISNULL(e.Art) OR NOT ISNULL(u.Ulcera) OR NOT ISNULL(di.ICD)) " & _
" AND f.pat_id = " & pid & _
" GROUP BY f.pat_id;"
myFrag rUlc, sql
If rUlc.BOF Then
 aktDC.ulcus = 2 ' kein Ulcus / 3 = nicht untersucht
Else
 aktDC.ulcus = rUlc!ulcus '0=oberflächlich, 1=tief
End If
' NaeUs As Integer ' 0=1 Jahr, 1=6 Monate, 2=3 Monate
' Infekt As Integer '0=nicht untersucht, 1=ja, 2=nein
' Deform As Integer '0=nein, 1=ja
' Hyperk As Integer '0=nein, 1=ja
' ZnUlcus As Integer '0=nein, 1=ja
' ZnAmput As Integer '0=nein, 1=ja

'"IF(fußdeform IN ('nein','keine','-',''),0,1) fußd,
myFrag rfuss, "SELECT IF(nae_us LIKE '%3%Mo%',2,IF(nae_us LIKE '%6%Mo%',1,0)) NaeUS, " & _
"IF(wundinfektion LIKE 'ja%',1,2) Infekt" & _
",fußdeform RLIKE 'schwarz|verfärb|hühner|rh?agad|gangl|ausschlag|ekzem|allergie|wunde|tophi|läsion|verkrust|infekt|extrem|weicht|fehl|verhärt|entzünd|ausgepr|hammer|verschieb|Zehenüberstand|Panaritium|ab(ge)?heil|verhärt|unter|über|fehlstand|aufger|mazer|geschw|Klump|ödem|beul|schmerz|röt|clavus|ch?a?[rt]c?h?ot?(-| |)[ -]?f|Chat- cot|Blase|Druckst|eingew|Amp(\.|u)|aput|orthese|klauen' fußd" & vbCrLf & _
",(hyper_mEin RLIKE 'D[12345]' OR hyper_mEin LIKE 'ja%') AND hyper_mEin<>fußdeform HypermEin, " & _
"NOT zn_ulcus IN ('nein','','-') AND zn_ulcus<>fußdeform ZnUlcus, " & _
"zn_amput NOT IN ('nein','','-') Amp, Fuß_ang " & _
"FROM fuss " & _
"WHERE pat_id = " & pid & _
" AND (zn_amput IN ('nein','ja') OR zn_amput<>fußdeform)" & _
" AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM fuss WHERE pat_id = " & pid & " AND (zn_amput IN ('nein','ja') OR zn_amput<>fußdeform) AND zeitpunkt BETWEEN qanf() AND qend() AND zeitpunkt > 20170717080000)"
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
myFrag rUlc, "SELECT pat_id FROM ulcus WHERE pat_id = " & pid & _
" AND (beläge RLIKE 'eitr|schmier' OR exsudat RLIKE 'gelb|braun|stark|viel|zünd|mehr|eitr|eiter' OR " & _
" geruch RLIKE 'ja|röt|mittel|stark|äßig|vorhand' OR wundrand RLIKE 'röt|wärm|infiz|infekt|rot' OR " & _
" (wundumgebung RLIKE 'röt|wärm|infiz|infekt|rot|verfärb' AND NOT wundumgebung RLIKE 'keine Röt')) "
If Not rUlc.BOF Then aktDC.Infekt = 1

If aktDC.Infekt = 1 Or aktDC.ulcus < 2 Then aktDC.NaeUs = 2

Dim raltu As New ADODB.Recordset
myFrag raltu, "SELECT icd FROM diagview d WHERE pat_id = " & pid & " AND icd RLIKE '^L89\.[234]' AND ((diagdatum < qanf() AND diagsicherheit<>'A') OR diagsicherheit='Z')"
If Not raltu.BOF Then aktDC.ZnUlcus = 1

Dim ranamp As New ADODB.Recordset
myFrag ranamp, "SELECT NOT ISNULL(amputation) AND amputation NOT IN ('n','-','','entfällt','/','--','nn','u','.-') AND amputation NOT LIKE '?%' AND amputation not RLIKE '^n[ .]' AND amputation not RLIKE 'nein' amp FROM anamnesebogen a WHERE pat_id = " & pid
If Not ranamp.EOF Then
 If Not ranamp.EOF And ranamp!amp = 1 Then aktDC.ZnAmput = 1
End If
 If aktDC.ab317 Then
  aktDC.x_wrstr = IIf(aktDC.Deform, "Fußdeformität, ", "") & _
   IIf(aktDC.Hyperk, "Hyperkeratose mit Einblutung, ", "") & _
   IIf(aktDC.ZnUlcus, "Z.n. Ulcus, ", "") & _
   IIf(aktDC.ZnAmput, "Z.n. Amputation, ", "") & _
   IIf(aktDC.Deform Or aktDC.Hyperk Or aktDC.ZnUlcus Or aktDC.ZnAmput, "  ja", "nein")
  aktDC.x_Ulcus = Switch(aktDC.ulcus = 0, "oberflächlich", aktDC.ulcus = 1, "tief", aktDC.ulcus = 2, "nein", 1, "nicht untersucht")
  aktDC.x_Wundinfektion = IIf(aktDC.Infekt = 0, "nicht untersucht", IIf(aktDC.Infekt = 1, "  ja", "nein"))
  If aktDC.ZnUlcus Or aktDC.ZnAmput Or aktDC.ulcus = 0 Or aktDC.ulcus = 1 Or aktDC.Infekt = 1 Then aktDC.NaeUs = 2 ' 20.7.25
  aktDC.x_Intervall = IIf(aktDC.NaeUs = 0, "jährlich", IIf(aktDC.NaeUs = 1, "alle 6 Monate", "alle 3 Monate oder häufiger"))
  If mitStr Then
   TabPr "•  Pulsstatus: ", aktDC.x_Pulsstatus
   TabPr "Sensibilität: ", aktDC.x_SensText
   TabPr "Weiteres Risiko für Ulcus: ", aktDC.x_wrstr
   TabPr "Ulcus: ", aktDC.x_Ulcus
   TabPr "Wundinfektion: ", aktDC.x_Wundinfektion
   TabPr "Intervall der künftigen Fußinspektionen: ", aktDC.x_Intervall
  End If ' mitstr
 End If ' aktdc.ab317 Then
 
 Dim raFa As New ADODB.Recordset
 Dim diIsql$
 diIsql = "SELECT ICD, DiagSeite, DiagDatum FROM diagview WHERE pat_id = " & pid & " AND gicd RLIKE '^M14.6|^T79\.|^L89\.|^T89\.|^T87.4|^Z44.1'" '  AND COALESCE(Dggel,0)=0 " ' M14.6 = Charcot, M14.2 = Arthorpathie
 Dim lddat As Date
 myFrag raFa, "SELECT MAX(bhfb) AS lddat FROM `faelle` WHERE pat_id = " & pid
 If Not raFa.BOF Then
  If Not IsNull(raFa!lddat) Then
   lddat = raFa!lddat
   diIsql = diIsql & " AND (obdauer <> 0 OR diagdatum >= " & DatFor_k(lddat) & ")"
  End If
 End If
 Set raDT = Nothing
' Call raDT.Open(diIsql, DBCn, adOpenDynamic, adLockReadOnly)
 myFrag raDT, diIsql
 Dim obUlc%, FSauf$
 obUlc = obPosi(Ulcera)
 If obUlc Or Not raDT.BOF Then
  aktDC.fußst = auff
  FSauf = "auffällig "
  If obUlc Then FSauf = FSauf & "(Anamnesebogen:Ulcera: " & Ulcera & ") "
  Do While Not raDT.EOF
   Dim diags$
   If Not IsNull(raDT!DiagSeite) Then diags = raDT!DiagSeite
   FSauf = FSauf & "ICD: " & raDT!ICD & diags & " (" & Format$(raDT!DiagDatum, "d.m.yy") & ") "
   raDT.Move 1
  Loop
  If aktDC.ab317 And mitStr Then TabPr "Fußstatus:", FSauf
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
'  END Enum

  If lies.obMySQL Then
   sql = "SELECT * FROM diagview WHERE (obdauer <> 0 OR (obdauer = 0 AND CONCAT(((month(diagdatum)+2) div 3)¡ YEAR(diagdatum)) = '" & ZQuart(DokuDat) & "')) AND diagsicherheit<>'A' AND pat_id = " & aktDC.Pat_ID ' AND COALESCE(Dggel,0)=0
   sql = REPLACE$(sql, "¡", ",")
  Else
   sql = "SELECT * FROM diagview WHERE (obdauer <> 0 OR (obdauer = 0 AND (int((month(diagdatum)+2) / 3) & YEAR(diagdatum)) = '" & ZQuart(DokuDat) & "')) AND diagsicherheit<>'A' AND pat_id = " & aktDC.Pat_ID ' AND COALESCE(Dggel,0)=0
   sql = REPLACE$(REPLACE$(sql, "CONCAT", ""), "¡", " & ")
  End If
  
'  SET rDT = aktDCb.OpenRecordset(sql) ' SET rDT = aktDCb.OpenRecordset("Diagnosen", dbOpenDynaset)
  Set raDT = Nothing
'  raDT.Open sql, DBCn, adOpenDynamic, adLockReadOnly
  myFrag raDT, sql, adOpenStatic
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
'  sql = "SELECT * FROM `" & stftdb & "`.jpg WHERE pat_id = " & CStr(PID) & " AND cstr(int(month(patdatum)/3)+1) & YEAR(patdatum) = """ & ZQuart(DokuDat) & """"
'  SET rafot = aktDCb.OpenRecordset(sql)
  
  Dim raFot As New ADODB.Recordset
#If False Then
'If lese.obMySQL THEN
' IF FtCn.State = 0 THEN
'  Call acon(FotT)
' END IF
' raFot.Open "SELECT * FROM Jpg WHERE pat_id = " & CStr(pid) & " AND CONCAT(floor(month(patdatum)+2) div 3,year(patdatum)) = """ & ZQuart(DokuDat) & """", FtCn, adOpenDynamic, adLockReadOnly
'Else
' raFot.Open "SELECT * FROM jpg WHERE pat_id = " & CStr(pid) & " AND cstr(int(month(patdatum)+2)/3) & YEAR(patdatum) = """ & ZQuart(DokuDat) & """", CStrAcc & StFtDB, adOpenDynamic, adLockReadOnly
'END IF
#Else
'  raFot.Open "SELECT dokname FROM `dokumente` WHERE pat_id = " & pid & " AND dokart = 'jpg' AND dokname LIKE '%Foto%' AND dokname LIKE '% WA%'", DBCn, adOpenDynamic, adLockReadOnly
  sql = "SELECT name FROM tmbrie WHERE pat_id = " & pid & " AND name LIKE '%Foto% WA%.jpg'"
  myFrag raFot, sql, adOpenStatic
#End If

  Dim maxDN As DFSNiveau, aktWA$, aktSei As DMPSeite
  Do While Not raFot.EOF
   Dim wa$, körperteil$, beschreibung$, p1&
   wa = left$(Trim$(Mid$(raFot!name, InStr(raFot!name, " WA") + 3)), 2)
   körperteil = Mid$(raFot!name, InStr(raFot!name, "Foto ") + 5)
   For p1 = 1 To Len(körperteil) - 1
    If IsNumeric(Mid$(körperteil, p1, 1)) And (Mid$(körperteil, p1 + 2, 1) = "." Or Mid$(körperteil, p1 + 1, 1) = ".") Then
     beschreibung = Mid$(körperteil, InStr(p1, körperteil, " ") + 1)
'     beschreibung = trim$(LEFT(beschreibung, InStr(beschreibung, "WA") - 1))
     Dim poswa%
     poswa = InStr(beschreibung, "WA ")
     If poswa = 0 Then poswa = InStr(beschreibung, "Wagner ")
     If poswa > 0 Then
      beschreibung = left$(beschreibung, poswa - 1)
     End If
     beschreibung = Trim$(beschreibung)
     körperteil = Trim$(left$(körperteil, p1 - 1))
     Exit For
    End If
   Next p1
   aktDN = stNichts
    If Not IsNull(wa) Then
     If Len(wa) > 0 Then
      Select Case left$(wa, 1)
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
  
'  IF aktDC.mwa = "" THEN
  If raDT.State = 0 Then
   Set raDT = Nothing
'  raDT.Open sql, DBCn, adOpenDynamic, adLockReadOnly
   myFrag raDT, sql, adOpenStatic
  End If
  If Not raDT.BOF Then
   raDT.MoveFirst
   aktSei = unbek
   Do While Not raDT.EOF
    If InStrB(raDT!ICD, "L89.") <> 0 Then
     aktWA = MAXvb(Mid$(raDT!ICD, 5, 1) - 1, 0) ' Korrektur 3.1.15 z.B. L89.09
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
  End If ' Not raDT.BOF THEN
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
'  END IF
  Dim DruckmSei$
  If aktDC.mSei = unbek Then aktDC.mSei = aktSei
  Select Case aktDC.mSei
   Case rE: DruckmSei = "rechts"
   Case li: DruckmSei = "links"
   Case gleich: DruckmSei = "gleich"
   Case unbek: DruckmSei = vNS
  End Select
  If aktDC.ab317 And mitStr Then TabPr "Schwerer betr. Fuß:", IIf(LenB(DruckmSei) = 0, DruckNStSeite, DruckmSei)
  If aktDC.mSei = gleich Or aktDC.mSei = unbek Then
   aktDC.mSei = aktDC.NStSeite
   If aktDC.mSei = gleich Or aktDC.mSei = unbek Then aktDC.mSei = rE ' meistens rechts
  End If
  If aktDC.mWA > "" Then
   If aktDC.ab317 And mitStr Then TabPr "Grad n. Wagner+Armstrong:", aktDC.mWA
  End If
  Set raFot = Nothing
'#If False THEN
'  Dim posi%
'  posi = InStr(Diagnosen, "Wagner")
'  Dim inf AS Boolean, isch AS Boolean, ArmSt$
'  inf = False
'  isch = False
'  IF diI("T79.") OR diI("T89") OR diI("T87.4") THEN inf = True
'  'If diI("M12") OR diI("M20") OR diI("M21") OR diI("M77.3") OR diI("M84") OR diI("M84.17") THEN def = True
'  IF (trp AND drp) OR (tlp AND dlp) OR (trp AND dlp) OR (tlp AND drp) THEN isch = True
'  ArmSt = "A"
'  IF inf THEN
'   ArmSt = "B"
'   IF isch THEN ArmSt = "D"
'  ElseIf isch THEN
'   ArmSt = "C"
'  END IF
'  IF posi <> 0 THEN
'   IF mitstr THEN TabPr "Grad nach Wagner und Armstrong: ", IIf(posi <> 0, Mid$(Diagnosen, posi + 7, 1) + ArmSt, "0A")
'  END IF
'#END IF
#If vorca2008 Then
  Set raDT = Nothing
'  Call raDT.Open("SELECT * FROM diagview WHERE pat_id = " & pid & " AND icd LIKE '" & "M14.6%" & "' AND diagsicherheit IN (""G"",""V"") AND COALESCE(Dggel,0)=0 ", DBCn, adOpenDynamic, adLockReadOnly)
  myFrag raDT, "SELECT * FROM diagview WHERE pat_id = " & pid & " AND gicd LIKE 'M14.6%'" ' AND COALESCE(Dggel,0)=0 "
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
   Else ' Not IsNull(raDT!DiagSeite) Then
    Seite = "  ja"
    aktDC.oap = gleich
   End If ' Not IsNull(raDT!DiagSeite) Then else
  Else ' Not raDT.BOF Then
   aktDC.oap = unbek ' heißt hier: keine
   Seite = "keine"
  End If ' Not raDT.BOF Then else
  If aktDC.ab317 And mitStr Then TabPr "Osteoarthropathie:", Seite
  Set raDT = Nothing
#End If ' vorca2008
 Else ' obUlc OR Not raDT.BOF THEN
'  If tfeld(Ulcera) = "" And tfeld(![Puls Atp]) = "" And tfeld(![Puls Adp]) = "" And tfeld(![Vibration IK]) = "" And tfeld(![Vibration Großzehe]) = "" And tfeld(![Kalt-Warm]) = "" Then
  If anaNdok Then
   aktDC.fußst = ndok
   FSauf = "" '"nicht dokumentiert"
  Else ' tfeld(Ulcera) = "" ...
   aktDC.fußst = unauff
   FSauf = "unauffällig"
  End If ' tfeld(Ulcera) = "" ... else
  If aktDC.ab317 And mitStr Then TabPr "Fußstatus:", FSauf
 End If ' obUlc OR Not raDT.BOF THEN else
 
 If aktDC.dtyp = "1" Then
  If mitStr Then
   If aktDC.PrRR <> "" Then TabPr "Blutdruck:", aktDC.PrRR
   TabPr "HbA1c:", aktDC.x_DMPHbA1c
   TabPr "Serum-Kreatinin:", DMPCrea
   TabPr "Urin-Albumin:", aktDC.x_DMPUAlb
  End If ' mitstr
 End If ' aktDC.dtyp = "1" Then
 
#If vorca2008 Then
' Dim oblaser AS Boolean
 Dim kontrah$
 aktDC.oblaser = True
 If Not obPosi(![Netzhaut gelasert]) Then
  aktDC.oblaser = False
'  rEi.Seek "=", PID
  Set raEi = Nothing
'  Call raEi.Open("SELECT * FROM `eintraege` WHERE pat_id = " & pid, DBCn, adOpenDynamic, adLockReadOnly)
  myFrag raEi, "SELECT * FROM `eintraege` WHERE pat_id = " & pid
  If Not raEi.BOF Then
   Do
    Dim StartPosi&
    If raEi.EOF Then Exit Do
    kontrah = REPLACE$(raEi!Inhalt, " ", "")
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
If aktDC.dtyp = "2" Then If mitStr Then TabPr "Lasertherapie:", IIf(aktDC.oblaser, "  ja", "nein")
#End If 'vorca2008
 
  If mitStr And aktDC.ab315 And Not aktDC.ab317 Then
   TabPr "•  Injektionsstellen:", aktDC.x_Inj
  End If
  
' 3. Reiter
 If diI("Z49", , fctQAnf(ZQuart(DokuDat))) Then
  aktDC.neuDial = True
  Set raDT = Nothing
 End If
 If diI("H54", , fctQAnf(ZQuart(DokuDat)), , "diagseite IN (' ','B')") Then
  aktDC.neuErbl = True
  Set raDT = Nothing
 End If
 If diI("Z44", , fctQAnf(ZQuart(DokuDat)), True) Then
  aktDC.neuAmp = True
  Set raDT = Nothing
 End If
 If diI("I21", , fctQAnf(ZQuart(DokuDat)), True) Then
  aktDC.neuMI = True
  Set raDT = Nothing
 End If
 If diI("I61 I62 I63 I64", , fctQAnf(ZQuart(DokuDat)), True) Then
  aktDC.neuApo = True
  Set raDT = Nothing
 End If
 
 If aktDC.dtyp = "1" Then If mitStr Then TabPr "Lasertherapie:", IIf(aktDC.oblaser, "  ja", "nein")
 If aktDC.ab315 Then
  If aktDC.FEn(8) Then ereig.AppVar Array("Nierenersatzth.", IIf(aktDC.FE(8), "(bek)", "(ICD " & aktDC.FEi(8) & " vor " & Format(aktDC.FEd(8), "d.m.yy") & ")"), ", ")
  If aktDC.FEn(10) Then ereig.AppVar Array("Erblindung", IIf(aktDC.FE(10), "(bek)", "(ICD " & aktDC.FEi(10) & " vor " & Format(aktDC.FEd(10), "d.m.yy") & ")"), ", ")
  If aktDC.FEn(13) Then ereig.AppVar Array("Amputation", IIf(aktDC.FE(13), "(bek)", "(ICD " & aktDC.FEi(13) & " vor " & Format(aktDC.FEd(13), "d.m.yy") & ")"), ", ")
  If aktDC.FEn(4) Then ereig.AppVar Array("Herzinfarkt", IIf(aktDC.FE(4), "(bek)", "(ICD " & aktDC.FEi(4) & " vor " & Format(aktDC.FEd(4), "d.m.yy") & ")"), ", ")
  If aktDC.FEn(5) Then ereig.AppVar Array("Schlaganfall", IIf(aktDC.FE(5), "(bek)", "(ICD " & aktDC.FEi(5) & " vor " & Format(aktDC.FEd(5), "d.m.yy") & ")"), ",")
  If ereig.Length = 0 Then ereig = "keine d.gen." Else ereig.Cut (ereig.Length - 2)
  aktDC.x_Ereig = ereig
  If mitStr Then TabPr "relev.Ereignisse: ", aktDC.x_Ereig
 End If ' aktdc.ab315 Then
 
 If Vgst > DokuDat - 92 Then ' Anamnese nur ein Quartal berücksichtigen
  aktDC.hypoZAn = 0
  If obPosi(![Fremde Hilfe pa]) Then
   If IsNumeric(![Fremde Hilfe pa]) Then
    aktDC.hypoZAn = ![Fremde Hilfe pa]
    aktDC.hypoZ = aktDC.hypoZAn
   End If
   aktDC.x_HypoStr = CStr(![Fremde Hilfe pa])
  Else
   aktDC.x_HypoStr = IIf(LenB(tfeld(![Fremde Hilfe pa])) = 0, "keine Angaben", "keine")
  End If
  If mitStr Then TabPr "Schw.Hypoglyk./12 Mon:", aktDC.x_HypoStr
 Else ' Vgst > DokuDat - 92 Then else
  myFrag rs, "SELECT COUNT(0) AS ct FROM `eintraege` WHERE pat_id = " & pid & " AND art = 'hypo' AND zeitpunkt >= " & DatFor_k(fctQAnf(ZQuart(DokuDat)))
  aktDC.hypoZKK = rs!ct
  Set rs = Nothing
  If IsNumeric(aktDC.dmpHypos) Then If aktDC.dmpHypos <> 0 Then aktDC.hypoZ = aktDC.dmpHypos
  aktDC.hypoZ = MAXvb(aktDC.hypoZ, aktDC.hypoZKK)
  If aktDC.hypoZ Then
   aktDC.x_HypoStr = aktDC.hypoZ
  Else ' aktDC.hypoZ else
   aktDC.x_HypoStr = "keine"
   myFrag rs, "SELECT GROUP_CONCAT(REPLACE(inhalt,'kommt ','komme ')) FROM `eintraege` WHERE pat_id = " & pid & " AND art = 'uzu' AND zeitpunkt >= " & DatFor_k(fctQAnf(ZQuart(DokuDat)))
   If Not rs.BOF Then
    aktDC.x_HypoStr = aktDC.x_HypoStr & " ('Unterzucker?': " & rs.Fields(0) & ")"
   End If ' not rs.bof
   Set rs = Nothing
  End If ' aktDC.hypoZ else
  If mitStr Then TabPr "Schw.Hypoglyk./Q.:", aktDC.x_HypoStr
 End If ' Vgst > DokuDat - 92 Then
 
 If Vgst > DokuDat - 92 Then ' Anamnese nur ein Quartal berücksichtigen
  Dim khsStr$
' Handlungsbedarf: Krankenhauseinweisungen
  aktDC.krZAn = 0
  If obPosi(![keto pa]) Then
   If IsNumeric(![keto pa]) Then
'   aktDC.krZAn = ![keto pa]
    If IsDate(![keto pa]) Then
     aktDC.krZAn = 1
    ElseIf IsNumeric(![keto pa]) Then
     aktDC.krZAn = ![keto pa]
    End If
    aktDC.krz = aktDC.krZAn
   End If
   khsStr = CStr(![keto pa])
  Else
   khsStr = IIf(LenB(tfeld(![Fremde Hilfe pa])) = 0, "keine Angaben", "keine")
  End If
  If mitStr Then TabPr "Krankenhausaufenthalte / Ketoazidosen /12 Mon:", khsStr
 Else ' Vgst > DokuDat - 92 Then ' Anamnese nur ein Quartal berücksichtigen
'  raKH.Open "SELECT * FROM kheinweis WHERE pat_id = " & pid & " AND zeitpunkt > " & DatFor_k(DokuDat - 92), DBCn, adOpenDynamic, adLockReadOnly
  myFrag raKH, "SELECT Zeitpunkt,Ziel,Diagnose FROM kheinweis WHERE pat_id = " & pid & " AND zeitpunkt > " & DatFor_k(DokuDat - 92), adOpenStatic
  If Not raKH.EOF Then
   aktDC.khew = True
   If mitStr Then TabPr "Krankenhauseinweisung", "(am " + Format$(raKH!Zeitpunkt, "dd/mm/yy") + IIf(tfeld(raKH!Ziel) <> "nkenhauseinweisung", " nach:" + raKH!Ziel, "") + _
             " mit: " + tfeld(raKH!Diagnose) + ")"
  End If
  Set raKH = Nothing
' nur ein Datensatz 2008
  myFrag rs, "SELECT COUNT(0) ct FROM `eintraege` WHERE pat_id = " & pid & " AND art = 'kra' AND zeitpunkt >= " & DatFor_k(fctQAnf(ZQuart(DokuDat)))
  aktDC.krZKK = rs!ct
  Set rs = Nothing
  If (IsNumeric(aktDC.dmpKhsA)) Then If aktDC.dmpKhsA <> 0 Then aktDC.krZKK = aktDC.dmpKhsA
  aktDC.krz = aktDC.krZKK
 End If ' Vgst > DokuDat - 92 Then Else ' Anamnese nur ein Quartal berücksichtigen
 
 
' 4. Reiter
 Dim TherArtAkt$
' TherArtAkt = TherArtAkt
' TherArtAkt = therart_erm(aktDC.Pat_id, 0)
 TherArtAkt = "Diät"
 On Error Resume Next
 TherArtAkt = myEFrag("SELECT therart FROM `therarten` tha WHERE pat_id=" & pid & " AND DATE(zp)<= " & Format(Now(), "yyyymmdd") & " ORDER BY zp DESC,mpnr DESC LIMIT 1").Fields(0)
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
' END SELECT
' IF aktDC.tart = csii THEN
 If aktDC.ab315 Then
  aktDC.x_Insulin = IIf(aktDC.obIns = adja Or aktDC.obAnal = adja, "  ja", "nein")
  If mitStr Then TabPr "•  Insulin oder -analoga:", aktDC.x_Insulin
 Else
  If mitStr Then TabPr "Insulin:", IIf(aktDC.obIns, IIf(aktDC.insz > 2, IIf(Not aktDC.ab315 And aktDC.VorM(5) = 1, "(unverändert) ", "  ") & "ja, intensiviert", IIf(Not aktDC.ab315 And aktDC.VorM(5) = 2, "(unverändert) ", "  ") & "ja, " & TherArtAkt), IIf(Not aktDC.ab315 And aktDC.VorM(5) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "nein")
  If mitStr Then TabPr "Insulin-Analoga:", IIf(aktDC.obAnal, IIf(aktDC.insz > 2, IIf(Not aktDC.ab315 And aktDC.VorM(6) = 1, "(unverändert) ", "  ") & "ja, intensiviert", IIf(Not aktDC.ab315 And aktDC.VorM(6) = 2, "(unverändert) ", "  ") & "ja, " & TherArtAkt), IIf(Not aktDC.ab315 And aktDC.VorM(6) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "nein")
 End If
 
 aktDC.obNI = (aktDC.eGFR < 45)
 
' IF mitStr THEN DmPStrS = DmPStrS & vbCrLf
 If aktDC.dtyp <> "1" Then
  If Not aktDC.ab1023 Then If mitStr Then TabPr "Glibenclamid:", IIf(aktDC.ab315, IIf(aktDC.obGlib, "  ja", IIf(aktDC.obNI, "Kontraindikation", "nein")), IIf(aktDC.obGlib, IIf(Not aktDC.ab315 And aktDC.VorM(0) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "  ja", IIf(Not aktDC.ab315 And aktDC.VorM(0) > 0, "(unverändert) ", "  ") & IIf(aktDC.obNI, "Kontraindikation", "nein")))
  If Not aktDC.obGlib And aktDC.obNI Then aktDC.obGlib = adki
  If aktDC.obmetf Then
   aktDC.x_Metformin = "  ja"
   If Not aktDC.ab315 And aktDC.VorM(1) = 0 And aktDC.obVorb Then aktDC.x_Metformin = "(unverändert)" & aktDC.x_Metformin
   If aktDC.obNI Then aktDC.x_Metformin = aktDC.x_Metformin & ": Sollte evtl. aufgrund Creatininerhöhung abgesetzt werden"
  Else
   If aktDC.obNI Then
    aktDC.obmetf = adki
    aktDC.x_Metformin = "Kontraindikation"
   Else ' aktDC.obNI Then
    aktDC.x_Metformin = "nein"
   End If ' aktDC.obNI Then else
   If Not aktDC.ab315 Then If aktDC.VorM(1) = 1 Or aktDC.VorM(1) = 2 Then aktDC.x_Metformin = "(unverändert)" & aktDC.x_Metformin
  End If ' aktDC.obmetf Then else
  If mitStr Then TabPr "Metformin:", aktDC.x_Metformin
  
  If aktDC.ab315 Then
   aktDC.x_Sonstige = IIf((aktDC.ab1023 And aktDC.obGlib = adja) Or aktDC.obGlucI = adja Or aktDC.obSHGlin = adja Or aktDC.obGlit = adja Or (Not aktDC.ab1023 And (aktDC.obSglt2 = adja Or aktDC.obGlp1 = adja)) Or aktDC.obDpp4 = adja Or aktDC.obSonstAD = adja, "  ja", "nein")
   If mitStr Then TabPr "Sonst.or.antidiab.Medik.:", aktDC.x_Sonstige
   If aktDC.ab1023 Then
    aktDC.x_Sglt2 = Switch(aktDC.obSglt2 = adja, "  ja", aktDC.obSglt2 = adnein, "nein", aktDC.obSglt2 = adki, "Kontraindikation")
    If mitStr Then TabPr "SGLT2-Inhibitor:", aktDC.x_Sglt2
    aktDC.x_Glp1 = Switch(aktDC.obGlp1 = adja, "  ja", aktDC.obGlp1 = adnein, "nein", aktDC.obGlp1 = adki, "Kontraindikation")
    If mitStr Then TabPr "GLP1-Rezeptoragonist:", aktDC.x_Glp1
    aktDC.x_OAK = Switch(aktDC.obOAK = adja, "  ja", aktDC.obOAK = adnein, "nein", aktDC.obOAK = adki, "Kontraindikation")
   End If ' aktdc.ab1023 Then
  Else ' aktdc.ab315 Then
   If aktDC.obGlit = adnein And diI("I50") Then aktDC.obGlit = adki
   If mitStr Then
    TabPr "Glucosidase-Inhib.:", IIf(aktDC.obGlucI, IIf(aktDC.VorM(2) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "  ja", IIf(aktDC.VorM(2) > 0, "(unverändert) ", "  ") & "nein")
    TabPr "Sonst. Sulf'hst./Glinide:", IIf(aktDC.obSHGlin, IIf(aktDC.VorM(3) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "  ja", IIf(aktDC.VorM(3) > 0, "(unverändert) ", "  ") & "nein")
    TabPr "Glitazone:", IIf(aktDC.obGlit = adja, IIf(aktDC.VorM(4) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "  ja", IIf(aktDC.VorM(4) > 0, "(unverändert) ", "  ") & IIf(aktDC.obGlit = adki, "Kontraindikation", "nein"))
    TabPr "Sonstige Antidiab.:", IIf(aktDC.obSonstAD = adja Or aktDC.obDpp4 = adja Or aktDC.obGlp1 = adja Or aktDC.obSglt2 = adja, IIf(aktDC.VorM(12) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "  ja", IIf(aktDC.VorM(12) > 0, "(unverändert) ", "  ") & "nein")
   End If ' mitstr
  End If ' aktdc.ab315 Then else
 End If ' aktDC.dtyp <> "1"
 aktDC.x_Thromb = IIf(aktDC.obThro, IIf(Not aktDC.ab315 And aktDC.VorM(9) = 1, "(unverändert) ", "") & "  ja", IIf(Not aktDC.ab315 And aktDC.VorM(9) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "nein") & IIf(Not aktDC.ab1023 And aktDC.obOAK, "; orale Antikoagul.", "")
 If mitStr Then TabPr "Thrombozytenhemmer:", aktDC.x_Thromb
 If aktDC.ab1023 Then
  If mitStr Then TabPr "Orale Antikoagulation:", aktDC.x_OAK
 Else ' aktdc.ab1023 Then
  If mitStr Then
   TabPr "ACE-Hemmer:", IIf(aktDC.obACEH, IIf(Not aktDC.ab315 And aktDC.VorM(11) = 1, "(unverändert) ", "  ") & "  ja", IIf(Not aktDC.ab315 And aktDC.VorM(11) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & IIf(aktDC.obAT1, "AT1", "nein"))
   TabPr "Betablocker:", IIf(aktDC.obBetabl, IIf(Not aktDC.ab315 And aktDC.VorM(10) = 1, "(unverändert) ", "  ") & "  ja", IIf(Not aktDC.ab315 And aktDC.VorM(10) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "nein")
'  DmPStrS = DmPStrS & vbCrLf
   TabPr "HMG-CoA-Red'-Hemmer:", IIf(aktDC.obHMG, IIf(Not aktDC.ab315 And aktDC.VorM(7) = 1, "(unverändert) ", "  ") & "  ja", IIf(Not aktDC.ab315 And aktDC.VorM(7) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "nein")
   If aktDC.ab315 Then
    TabPr "Diuretika:", IIf(aktDC.obDiur, IIf(Not aktDC.ab315 And aktDC.VorM(13) = 1, "(unverändert) ", "  ") & "  ja", IIf(Not aktDC.ab315 And aktDC.VorM(13) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "nein")
   Else ' aktdc.ab315 Then
    TabPr "Antihypertensiva:", IIf(aktDC.obAntihyp, IIf(Not aktDC.ab315 And aktDC.VorM(8) = 1, "(unverändert) ", "  ") & "  ja", IIf(Not aktDC.ab315 And aktDC.VorM(8) = 0 And aktDC.obVorb, "(unverändert) ", "  ") & "nein")
   End If ' aktdc.ab315 Then else
  End If ' mitstr
 End If ' aktdc.ab1023

' 5. Reiter
' IF mitStr THEN DmPStrS = DmPStrS & vbCrLf
 If Not aktDC.ab315 Then If mitStr Then TabPr "•  Schulungen vor DMP:", IIf(aktDC.SW(0) <> "", " D.m.: " & aktDC.SW(0), "") & IIf(aktDC.SW(1) <> 0, " Hypert: " & aktDC.SW(1), "") & IIf(aktDC.SW(0) = "" And aktDC.SW(1) = "", " keine", "")
 
' Schulungen empfohlen
' SET raLau = Nothing
' raLau.Open "SELECT zeitpunkt,CAST(wert AS DECIMAL(10,1)) wert FROM (" & lsql & ") AS innen WHERE abkü RLIKE '^HBA[1C]' AND CAST(wert AS decimal) < 22 ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly ' wegen falsch eingetragener Fremdlabore gestrichen: AND einheit = '%'
' SET raLau = hollabor(Pat_id, "^HBA[1C]", 0, 22)
 If lwZahl Then
  alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
  Labs = LabPat(LA_HbA1c, Pat_ID)
  Dim zpd1$, zpdL$, zpr1$, zprL$
' IF Not raLau.BOF THEN
  If Labs.Abkü <> "" Then
   If Labs.WertSg <> vNS Then
    pos = InStr(Labs.WertSg, ">") ' Pat_id 744
    If pos > 0 Then Labs.WertSg = Mid$(Labs.WertSg, pos + 1)
    If IsNumeric(Labs.WertSg) Then
     If Labs.WertSg >= 8# Then
      DMSchulz = SchulzBest(CStr(Pat_ID), zpd1, zpdL, DokuDat - 2 * 365)
      If ((aktDC.tart = OAD Or aktDC.tart = offen Or aktDC.tart = Diät) And DMSchulz < 4) Or DMSchulz < 6 Then
       If zplschul(Pat_ID) < DokuDat - 92 Then
'    Debug.Print Pat_id
        If aktDC.GebDat > DokuDat - 365 * 83 Then  ' Schulung nicht mehr unbedingt über 83 Jahre empfehlen
         aktDC.SE(0) = DokuDat ' SE(2) Schulung empfohlen
        End If ' aktDC.GebDat > DokuDat - 365
       End If ' zplschul(Pat_ID) < DokuDat - 92 Then
      End If ' ((aktDC.tart = OAD Or aktDC.tart = offen
     End If ' Labs.WertSg >= 8# Then
    End If ' IsNumeric(Labs.WertSg) Then
   End If ' Labs.WertSg <> vNS Then
  End If ' Labs.Abkü <> "" Then
 End If ' lwzahl
 If aktDC.dmpDMSchulEmpf = "j" Then
  On Error Resume Next
  aktDC.SE(0) = CDate(myEFrag("SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = " & CStr(pid) & " AND art = 'dmperg' AND instr(inhalt,'D.m.-Schul.empf.: j')").Fields(0))
  On Error GoTo fehler
  If aktDC.SE(0) = 0 Then aktDC.SE(0) = DokuDat - 1
 ElseIf aktDC.dmpDMSchulEmpf = "n" Then
  aktDC.SE(0) = 0
 End If
 
 If aktDC.dmpHypertSchulEmpf = "j" Then
  aktDC.SE(1) = 0
  On Error Resume Next
  aktDC.SE(1) = CDate(myEFrag("SELECT MAX(zeitpunkt) FROM `eintraege` WHERE pat_id = " & CStr(pid) & " AND art = 'dmperg' AND instr(inhalt,'Hypert-Schul.empf.: j')").Fields(0))
  On Error GoTo fehler
  If aktDC.SE(1) = 0 Then aktDC.SE(1) = DokuDat - 1
 ElseIf aktDC.dmpHypertSchulEmpf = "n" Then
  aktDC.SE(1) = 0
 End If
 
 If Not aktDC.ab315 Then
  If obPosi(![Jahr letzte Diabetesschulung]) Then
   Dim OrtS$
   If Not IsNull(![Ort Schulung]) Then OrtS = " (" + ![Ort Schulung] + ")"
   aktDC.x_SchulStr = ![Jahr letzte Diabetesschulung] + OrtS
  Else ' obPosi(![Jahr letzte Diabetesschulung]) Then
   aktDC.x_SchulStr = "keine"
  End If ' obPosi(![Jahr letzte Diabetesschulung]) Then else
  If mitStr Then
   TabPr "D.m.-Schul.empfohl:", IIf(aktDC.SE(0) > 0, " D.m. " & Format$(aktDC.SE(0), "d.m.yy"), "") & IIf(aktDC.SE(1) > 0, " Hypert: " & Format$(aktDC.SE(1), "d.m.yy"), "") & IIf(aktDC.SE(2) > 0, " keine: " & Format$(aktDC.SE(2), "d.m.yy"), "")
   TabPr "DMP-Schulung D.m.:", IIf(aktDC.Ds(0) > 0, " ja: " & Format$(aktDC.Ds(0), "d.m.yy"), "") & IIf(aktDC.Ds(1) > 0, " nein: " & Format$(aktDC.Ds(1), "d.m.yy"), "") & IIf(aktDC.Ds(2) > 0, " nicht mögl: " & Format$(aktDC.Ds(2), "d.m.yy"), "")
   TabPr "DMP-Schulung Hypt:", IIf(aktDC.HS(0) > 0, " ja: " & Format$(aktDC.HS(0), "d.m.yy"), "") & IIf(aktDC.HS(1) > 0, " nein: " & Format$(aktDC.HS(1), "d.m.yy"), "") & IIf(aktDC.HS(2) > 0, " nicht mögl: " & Format$(aktDC.HS(2), "d.m.yy"), "")
   If aktDC.kSE <> 0 Then TabPr " Zuletzt k.Sch.empf:", Format$(aktDC.kSE, "d.m.yy")
   TabPr "Schulung vor Zuweisung:", aktDC.x_SchulStr
  End If ' mitStr Then
 End If ' Not aktdc.ab315 Then
 
 aktDC.x_SchulEmpf = IIf(aktDC.dmpDMSchulEmpf = "j", "Diabetes, ", "") & IIf(aktDC.dmpHypertSchulEmpf = "j", "Hypertonie, ", "") & IIf(aktDC.dmpDMSchulEmpf <> "j" And aktDC.dmpHypertSchulEmpf <> "j", "keine, ", "")
 aktDC.x_SchulEmpf = left$(aktDC.x_SchulEmpf, Len(aktDC.x_SchulEmpf) - 2)
 If mitStr Then TabPr "•  Schulungen empfohlen: ", aktDC.x_SchulEmpf
 
 Select Case aktDC.dmpDMSchulWahrg
  Case "j"
   aktDC.x_SchulWahrg = "  ja"
  Case "n"
   aktDC.x_SchulWahrg = "nein"
  Case "n mögl"
   aktDC.x_SchulWahrg = "nicht möglich"
  Case Else ' "u", "n empf"
   aktDC.x_SchulWahrg = "b.d.letzten Doku n.empf."
 End Select
 If mitStr Then TabPr "Diabetesschulung   wahrg.: ", aktDC.x_SchulWahrg
 
 Select Case aktDC.dmpHypertSchulWahrg
  Case "j"
   aktDC.x_HSchulWahr = "  ja"
  Case "n"
   aktDC.x_HSchulWahr = "nein"
  Case "n mögl"
   aktDC.x_HSchulWahr = "nicht möglich"
  Case Else ' "u", "n empf"
   aktDC.x_HSchulWahr = "b.d.letzten Doku n.empf."
 End Select
 If mitStr Then TabPr "Hypertonieschulung wahrg.: ", aktDC.x_HSchulWahr
 
 Dim sqls$
 sqls$ = _
 "SELECT `leistungen`.pat_id, `leistungen`.zeitpunkt, UCASE(`leistungen`.leistung) lst" & _
 " FROM `leistungen` INNER JOIN `faelle` ON `faelle`.pat_id = `leistungen`.pat_id AND `faelle`.quartal = cstr(int(month(`leistungen`.zeitpunkt)+2) / 3) & YEAR(`leistungen`.zeitpunkt) " & _
 " WHERE `faelle`.pat_id = " & CStr(pid)
 If lies.obMySQL Then sqls = REPLACE$(REPLACE$(REPLACE$(sqls, "int(", "floor("), "cstr", ""), " / ", " div ")
' SET lpp = aktDCb.OpenRecordset(sqls, dbOpenDynaset)
' Call lapp.Open(sqls, DBCn, adOpenDynamic, adLockReadOnly)
 myFrag lapp, sqls
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
 End If ' Not lapp.BOF Then
 
 DMSchulz = SchulzBest(CStr(pid), zpd1, zpdL)
 aktDC.x_SchulStr = IIf(DMSchulz > 1, CStr(DMSchulz) + " (Diabetes: " & zpd1 & " - " & zpdL & ")", "") + IIf(HSchulz > 1, CStr(HSchulz) + " (Hypertonie: " & zpr1 & " - " & zprL & ")", "")
 If aktDC.x_SchulStr = "" Then aktDC.x_SchulStr = "keine"
 If mitStr Then TabPr "Bish.Schulg.b.uns (ges.):", aktDC.x_SchulStr
'  DmPStrS = DmPStrS & vbCrLf
 
 
' 6. Reiter
 If aktDC.ab315 Then
  aktDC.x_InfoAng = ""
  If aktDC.dmpKKTabakEmpf = "j" Or aktDC.dmpKKkTrainEmpf = "j" Or aktDC.dmpKKErnEmpf = "j" Then
   If aktDC.dmpKKTabakEmpf = "j" Then aktDC.x_InfoAng = "Tabakverzicht, "
   If aktDC.dmpKKErnEmpf = "j" Then aktDC.x_InfoAng = aktDC.x_InfoAng & "Ernährungsberatung, "
   If aktDC.dmpKKkTrainEmpf = "j" Then aktDC.x_InfoAng = aktDC.x_InfoAng & "körperl.Training, "
   aktDC.x_InfoAng = left$(aktDC.x_InfoAng, Len(aktDC.x_InfoAng) - 2)
  Else
   aktDC.x_InfoAng = "nichts angegeben"
  End If
  If mitStr Then TabPr "•  Info'angeb.KK gewü.: ", aktDC.x_InfoAng
 Else ' aktdc.ab315 Then
  If aktDC.dtyp = "2" Then
   If mitStr Then TabPr "  Aufgabe Tabak empf.:", IIf(aktDC.Tabak, "  ja", "nein")
   If Größe <> 0 Then
''   IF mitstr THEN TabPr "Ernährungsber. empf.:", IIf(CDbl(replace$(!`bmi`, ".", ",")) > 24.9, "  ja", "nein")
'    Dim bmi!
'    bmi = !Gewicht * IIf(!Gewicht < 3, 100, 1) / !Größe / !Größe '* 10000
'    Do
'     IF bmi = 0 THEN Exit Do
'     IF bmi > 8 THEN Exit Do
'     bmi = bmi * 10
'    Loop
'    IF bmi > 24.9 THEN aktDC.ernb = True
    If aktDC.bmi >= 25 Then aktDC.ernb = True
    If mitStr Then TabPr "Ernährungsber. empf.:", IIf(aktDC.ernb, "  ja", "nein")
   End If ' Größe <> 0 Then
  End If ' aktDC.dtyp = "2" Then
 End If ' aktdc.ab315 Then else
 
 Dim grenze#
 If PAlter >= 75 Then
  grenze = 7.5
 ElseIf PAlter >= 70 Then
  grenze = 7
 ElseIf PAlter >= 65 Then
  grenze = 6.8
 Else
  grenze = 6.5
 End If
 aktDC.hbEmpf = IIf(aktDC.bekHb > grenze, senken, halten)
 If IsNumeric(aktDC.dmpHbA1cZiel) Then
   If aktDC.bekHb <= CDbl(aktDC.dmpHbA1cZiel) Then
    aktDC.hbEmpf = halten
   Else
    aktDC.hbEmpf = senken
   End If
 End If
 If aktDC.ab315 Then
  aktDC.x_HZiel = IIf(aktDC.hbEmpf = halten, "erreicht", "noch nicht erreicht")
  If mitStr Then TabPr "HbA1c-Zielwert: ", aktDC.x_HZiel
 Else
  aktDC.x_HZiel = IIf(aktDC.hbEmpf = senken, "senken", "halten")
  If mitStr Then TabPr "Zielvereinb. HbA1c:", aktDC.x_HZiel
  Dim rrEmpfehlung$
  rrEmpfehlung = rrEmpf(aktDC.RRsyst, aktDC.RRdiast, CStr(pid))
  aktDC.rrEmpf = IIf(rrEmpfehlung = "halten", halten, senken)
  If mitStr Then TabPr "Zielvereinb. Blutdruck:", rrEmpfehlung ' manuell auszufüllen
 End If ' aktdc.ab315 Then else
 
 If aktDC.ab315 Then
  aktDC.x_BhdF = IIf(aktDC.fußst = auff Or aktDC.dmpUewFuss = "j", "  ja", "nein")
  aktDC.x_BhdD = IIf(aktDC.dmpEinwDM = "j", "  ja", "nein")
  If mitStr Then
   TabPr "Bhdlg.Fußeinr.: ", aktDC.x_BhdF
   TabPr "Diab'bez.stat.Einweisung: ", aktDC.x_BhdD
  End If ' mitStr
 End If ' aktdc.ab315 Then
 
 If aktDC.dtyp = "1" Or Not aktDC.ab1023 Then
  Dim AugUDat As Date
  Dim rBr As New ADODB.Recordset
  aktDC.x_Aug = vNS
' SET rdo = aktDCb.OpenRecordset("SELECT * FROM `" + QMdbAkt + "`.`dokumente` WHERE pat_id = " & CStr(Pat_id) + " AND zeitpunkt > DokuDat - 550 ORDER BY zeitpunkt DESC", dbOpenDynaset) ' 1,5 Jahre
' rBr.Open "SELECT Quelldatum,Name,Zeitpunkt FROM `tmbrie` WHERE pat_id = " & CStr(aktDC.Pat_id) & " AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " AND name LIKE '%augen%' ORDER BY quelldatum DESC, zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
  myFrag rBr, "SELECT Quelldatum,Name,Zeitpunkt FROM `tmbrie` WHERE pat_id = " & CStr(aktDC.Pat_ID) & " AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " AND name LIKE '%augen%' ORDER BY quelldatum DESC, zeitpunkt DESC"
  If Not rBr.BOF Then
   AugUDat = rBr!Quelldatum
   If AugUDat = CDate(0) Then AugUDat = DatInStr(rBr!name, Year(rBr!Zeitpunkt))
   aktDC.x_Aug = "Befund vo" + IIf(AugUDat = CDate(0), "n " + Format$(rBr!Zeitpunkt, "yyyy"), "m " + Format$(AugUDat, "dd/mm/yy")) + " vorliegend"
   If AugUDat = CDate(0) Then AugUDat = rBr!Zeitpunkt ' dann war die letzte Untersuchung vermutlich frühestens am 1.1. des Scan-Jahres
  Else ' Not rBr.BOF Then
   Dim rAEin As New ADODB.Recordset
'   rAEin.Open "SELECT ZeitPunkt FROM eintraege WHERE pat_id = " & CStr(aktDC.Pat_id) & " AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " AND art='aug' AND (NOT inhalt RLIKE 'nie|nicht|lange|länger|(kein|will|wird|macht).*Termin' OR inhalt RLIKE 'Veränd|in Ordnung') ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
   myFrag rAEin, "SELECT ZeitPunkt FROM eintraege WHERE pat_id = " & CStr(aktDC.Pat_ID) & " AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " AND art='aug' AND (NOT inhalt RLIKE 'nie|nicht|lange|länger|(kein|will|wird|macht).*Termin' OR inhalt RLIKE 'Veränd|in Ordnung') ORDER BY zeitpunkt DESC"
   If Not rAEin.BOF Then
    AugUDat = rAEin!Zeitpunkt
   End If
  End If ' Not rBr.BOF Then Else
' Dim rEin AS DAO.Recordset
  If lies.obMySQL Then
'  sql = "SELECT * FROM `eintraege` WHERE pat_id = " & CStr(pid) & " AND ( art = 'aug' OR ((art IN (" & artspezG & ") AND ((inhalt LIKE ""%augenb%"" AND NOT inhalt LIKE ""%augenbl%"" AND NOT inhalt LIKE ""%augen"") OR (inhalt LIKE ""%augenarzt%"" OR inhalt LIKE ""%augenärzt%"") OR (inhalt LIKE ""% aa%"" AND NOT inhalt LIKE ""% aag%""))) OR (art = ""aa"" OR art = ""augen"")))" & " AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " ORDER BY zeitpunkt DESC LIMIT 3"
   sql = "SELECT * FROM `eintraege` WHERE pat_id = " & CStr(pid) & " AND ( art = 'aug' OR ((art IN (" & artspezG & ") AND ((inhalt LIKE ""%augenb%"" AND NOT inhalt LIKE ""%augenbl%"" AND NOT inhalt LIKE ""%augen"") OR (inhalt LIKE ""%augenarzt%"" OR inhalt LIKE ""%augenärzt%"") OR (inhalt LIKE ""% aa%"" AND NOT inhalt LIKE ""% aag%""))) OR (art = ""aa"" OR art = ""augen"")))" & " AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " ORDER BY zeitpunkt DESC LIMIT 3"
  Else
'  sql = "SELECT top 3 * FROM `eintraege` WHERE pat_id = " & CStr(pid) & " AND ( art = 'aug' OR ((art IN (" & artspezG & ") AND (inhalt LIKE ""%augenb`!l`%"" OR inhalt LIKE ""%augen`aä`rzt%"" OR inhalt LIKE ""% aa`!g`%"")) OR (art = ""aa"" OR art = ""augen"")))" + " AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " ORDER BY zeitpunkt DESC"
   sql = "SELECT top 3 * FROM `eintraege` WHERE pat_id = " & CStr(pid) & " AND ( art = 'aug' OR ((art IN (" & artspezG & ") AND (inhalt LIKE ""%augenb`!l`%"" OR inhalt LIKE ""%augen`aä`rzt%"" OR inhalt LIKE ""% aa`!g`%"")) OR (art = ""aa"" OR art = ""augen"")))" + " AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " ORDER BY zeitpunkt DESC"
  End If ' lies.obMySQL Then
' SET rEin = aktDCb.OpenRecordset(sql)
  Set raEi = Nothing
' raEi.Open sql, DBCn, adOpenDynamic, adLockReadOnly
  myFrag raEi, sql
  Do While Not raEi.EOF
   aktDC.x_Aug = aktDC.x_Aug & IIf(LenB(aktDC.x_Aug) <> 0, vbCrLf & vbTab, "") & "Eintrag am " & Format$(raEi!Zeitpunkt, "d/m/yy") & ":" + vbTab + raEi!Inhalt
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
'  fiabfr = "SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, form_abk FROM formulari WHERE pat_id=" & Pat_id & " AND form_abk='uew' AND feld ='Ueberweisung_an' AND feldinh ='Augenheilkunde' AND zeitpunkt > " & DatFor_k(DokuDat - 550) & " ORDER BY zeitpunkt DESC"
' ohne FORCE INDEX geht es manchmal ...
  fiabfr = "SELECT zeitpunkt " & vbCrLf & _
           "FROM forminhfeld fi " & vbCrLf & _
           "LEFT JOIN forminhkopf k FORCE INDEX (NamenFormInhKopf) ON k.foid=fi.foid" & vbCrLf & _
           "LEFT JOIN forminhfeld ff ON ff.foid=k.foid" & vbCrLf & _
           "LEFT JOIN formulare f ON f.FormID=k.form_id" & vbCrLf & _
           "LEFT JOIN forminhaltfeld fif ON fif.FeldVW=fi.feldvw" & vbCrLf & _
           "LEFT JOIN forminhaltfeldinh fifi ON fifi.FeldInhVW=fi.FeldInhVW" & vbCrLf & _
           "WHERE pat_id=" & Pat_ID & " AND form_abk='uew' AND feld='ueberweisung_an' AND feldinh='Augenheilkunde' AND zeitpunkt>" & DatFor_k(DokuDat - 550) & vbCrLf & _
           "ORDER BY zeitpunkt DESC LIMIT 1;"
  myFrag uebw, fiabfr
' SET uebw = aktDCb.OpenRecordset(, dbOpenDynaset)
  If Not uebw.BOF Then
   If AugUDat = CDate(0) Or uebw!Zeitpunkt > AugUDat Then
    aktDC.x_Aug = aktDC.x_Aug + IIf(LenB(aktDC.x_Aug) = 0, vNS, ", ") + "Untersuchung veranlasst am " + Format$(uebw!Zeitpunkt, "dd/mm/yy")
   End If
  End If
  If aktDC.x_Aug = "" And AspZul <> "" Then
   aktDC.x_Aug = AspZul
   If AspBef <> "" Then
    aktDC.x_Aug = aktDC.x_Aug + " (Befund: " + AspBef + ")"
   End If
  End If
  If AugUDat > DokuDat - 500 Then aktDC.aug = durchg Else aktDC.aug = veranl
  Dim crpos1%, crpos2%
  crpos1 = InStr(aktDC.x_Aug, vbCrLf)
  crpos2 = InStr(crpos1 + 1, aktDC.x_Aug, vbCrLf)
  If crpos1 <> 0 Then
   If crpos2 <> 0 Then
    aktDC.x_Aug = REPLACE$(left$(aktDC.x_Aug, crpos2), vbTab, ", ")
   Else
    aktDC.x_Aug = left$(aktDC.x_Aug, crpos1)
   End If
   aktDC.x_Aug = REPLACE$(aktDC.x_Aug, vbCr, "")
   aktDC.x_Aug = REPLACE$(aktDC.x_Aug, vbLf, "")
  End If ' crpos1 <> 0 Then
  If mitStr Then TabPr "Augenuntersuchung:", aktDC.x_Aug
 End If ' aktDC.dtyp = "1" Or Not aktdc.ab1023 Then
 aktDC.x_EmpfItv = IIf(aktDC.hbEmpf = halten And (Not obPosi(UzuPm) Or InStrB(UzuPm, "selten") <> 0), "viertel- oder halbjährlich", "vierteljährlich")
 If mitStr Then TabPr "Empf. Dok'intervall:", aktDC.x_EmpfItv
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
'If Err.Number = -2147467259 OR Err.Number = -2147217887 THEN ' Server has gone away / ungültiger Eigenschaftswert
 Call DBCnOpen
 Resume nochmal
End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DMPString/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' DMPString

Function zuJahr$(ByVal ursp$, Optional dokzt As Date)
 Dim buch$, pos&
 If ursp = "?" Or ursp = "" Then
  If dokzt = 0 Then dokzt = Now
  zuJahr = Year(dokzt)
 Else
  ursp = Trim$(REPLACE$(REPLACE$(REPLACE$(ursp, "ca.", ""), "ca", ""), "seit", ""))
  pos = InStr(ursp, "/")
  If pos > 0 Then
   ursp = Mid$(ursp, pos + 1)
   buch = Mid$(ursp, 3)
   If Not IsNumeric(buch) Then
    buch = left$(ursp, 1)
    If Not IsNumeric(buch) Then zuJahr = vNS: Exit Function
    If buch > 5 Then
     ursp = "19" & ursp
    Else
     ursp = "20" & ursp
    End If
   End If
  End If
  pos = InStr(ursp, " ")
  If pos > 0 Then
   If IsNumeric(left$(ursp, 1)) Then
    ursp = left$(ursp, pos - 1)
   Else
    ursp = Mid$(ursp, pos + 1)
   End If
  End If
  pos = InStr(ursp, "-")
  If pos > 0 Then
   ursp = left$(ursp, pos - 1)
  End If
  pos = InStr(ursp, "a")
  If pos > 0 Then
   ursp = left$(ursp, pos - 1)
  End If
  pos = InStr(ursp, "J")
  If pos > 0 Then
   ursp = left$(ursp, pos - 1)
  End If
  If InStrB(ursp, ".") > 0 Then
   If IsDate(ursp) Then
    zuJahr = Year(ursp)
   End If
  Else
   If Not IsNumeric(ursp) Then
    If dokzt = 0 Then dokzt = Now
    zuJahr = Year(dokzt)
   ElseIf ursp < 50 Then
    If dokzt = 0 Then dokzt = Now
    zuJahr = Year(dokzt - ursp * 365)
   ElseIf ursp < 100 Then
    zuJahr = "19" & ursp
   Else
    zuJahr = ursp
   End If
  End If
 End If
End Function ' zuJahr

#If False Then
' in TherapieArtEinzelnFestlegen, testTherArt
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
#End If

Function ImportFolderHerricht()
 Dim fld As Folder, Fil As File, FPath$
 On Error GoTo fehler
 Call VerzPrüf(hVerz)
 Set fld = FSO.GetFolder(hVerz)
 For Each Fil In fld.Files
  If Right$(Fil.name, 4) = ".BDT" Then
   FPath = Fil.path
   On Error Resume Next
   Name FPath As REPLACE$(FPath, ".BDT", ".txt")
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ImportFolderHerricht/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ImportFolderHerricht()
'Function LeistungsExport0()
' Dim op$, BDT As New BDTSchreib
' ON Error GoTo fehler
'' Call ImportFolderHerricht
'' LeistBDT = hVerz + "LEIST " + Format$(Now, "dd/mm/yy HH.MM") + ".BDT"
'' Open LeistBDT For Output AS #310
'  IF Not BDT.Start(hVerz, "Leist") THEN
'   Exit Function
'  END IF
'  Print #310, "01380000020"
'  Print #310, "014810000082"
'  Print #310, "0169100" & KVNR
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
'  Print #310, "0160201" & KVNR
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
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LeistungsExport0/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' LeistungsExport0

' in doDiagnosenExport und Leistungsexport1
Function FallExport(BDT As BDTSchreib, Pat_ID&, Datu As Date, ByRef Lanr&, Optional nurKasse%, Optional Arztnr&)
 Dim rFa As New ADODB.Recordset
 Dim op$
 On Error GoTo fehler
 myFrag rFa, "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_ID & " AND bhfb<=" & DatFor_k(Datu) & " AND (bhfe1 >= " & DatFor_k(Datu) & " OR bhfe1=18991230) " & IIf(nurKasse, " AND schgr <> '90' ", "") & " ORDER BY bhfe1"
 If rFa.BOF Then ' wenn kein Fall in die Zeitschiene paßt, dann den letzten vor dem Datum nehmen
  Set rFa = Nothing
  myFrag rFa, "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_ID & " AND bhfb<=" & DatFor_k(Datu) & IIf(nurKasse, " AND schgr <> '90' ", "") & " ORDER BY bhfe1 DESC"
  If rFa.BOF Then ' wenn kein Fall in die Zeitschiene paßt, dann den letzten nehmen
   Set rFa = Nothing
   myFrag rFa, "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_ID & IIf(nurKasse, " AND schgr <> '90' ", "") & " ORDER BY bhfe1 DESC"
   If rFa.BOF Then
    Exit Function
   End If
  End If ' rFa.BOF THEN
 End If ' rFa.BOF THEN
 If Arztnr <> 0 And rFa!lanrid <> Arztnr Then Exit Function
 FallExport = 1
 Lanr = IIf(IsNull(rFa!Lanr), 0, rFa!Lanr)
 Call BDT.Satzart(IIf(rFa!SchGr = "90", "0190", "0102")) ' 80000 Satzidentifikation
 Call BDT.PatID(rFa!Pat_ID)
     BDT.SAdd "3635", "TM#" & rFa!Lanr
     BDT.SAdd "3636", "TM#" & BSNR
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
     End If ' rFa!SchGr <> "90" THEN
     If Not IsNull(rFa!GebOr) And rFa!GebOr <> vNS Then ' bei Privaten
      BDT.SAdd "4121", rFa!GebOr
     End If
     If Not IsNull(rFa!AbrGb) And rFa!AbrGb <> vNS Then ' bei Privaten
      BDT.SAdd "4122", rFa!AbrGb
     End If ' NOT ISNULL(rFa!GebOr) THEN ' bei Privaten
     BDT.SAdd "4144", "TM#" & IIf(IsNull(rFa!TMFNr), Space$(11), rFa!TMFNr)
     BDT.DAdd "4150", rFa!BhFB
     BDT.DAdd "4151", rFa!BhFE1
'     op = format$(3 + 4 + 8, "000") + "4152" + IIf(rFa!BhFE2 = 0, "00000000", format$(rFa!BhFE2, "ddmmyyyy"))
'     Print #310, ZSU(op)
     Dim Üw$
     BDT.SAdd "4218", "TM#" & rFa!ÜbWVLANR & "#" & rFa!ÜbWVBSNR & "#" & rFa!ÜbWVKVNR
     Üw = rFa!Übwr ' IIf(ISNULL(rFa!AndÜw), vNS, rFa!AndÜw)
' 24.9.11: nicht nötig
'     IF Üw <> vNS THEN
'      BDT.SAdd "4219", Üw
'     END IF
'     IF NOT ISNULL(rFa!ÜWZiel) AND rFa!ÜWZiel <> "" THEN
'      BDT.SAdd "4220", rFa!ÜWZiel
'     END IF
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FallExport/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' FallExport

Function LeistungsExport1(BDT As BDTSchreib, Pat_ID&, Leist$, Datu As Date, Optional QUZeit$, Optional nurKasse%, Optional Arztnr&)
 Dim rNa As New ADODB.Recordset, Lanr&
 myFrag rNa, "SELECT * FROM `namen` WHERE pat_id = " & Pat_ID
 Lese.Ausgeb "Trage Leistung " & Leist & " für Pat_ID " & Pat_ID & " (" & rNa!Nachname & ", " & rNa!Vorname & ") mit Datum " & Format$(Datu, "dd.mm.yyyy") & " ein.", True
 Open pVerz & LEDatei For Append As #347
 Print #347, Pat_ID & ": " & rNa!Nachname & ", " & rNa!Vorname & ", " & Format(Datu, "dd.mm.yyyy")
 Close #347
 If FallExport(BDT, Pat_ID, Datu, Lanr, nurKasse, Arztnr) <> 1 Then Exit Function
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
     BDT.SAdd "3636", "TM#" & BSNR
     BDT.SAdd "5001", Leist
     BDT.SAdd "5098", BSNR
     BDT.SAdd "5099", CStr(Lanr)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LeistungsExport1/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LeistungsExport1

#If False Then
Function LeistungsExport1a(BDT As BDTSchreib, Pat_ID&, Leist$, Datu As Date, Optional QUZeit$, Optional nurKasse%, Optional Arztnr%)
' Dim q AS DAO.Recordset
 Dim rFa As New ADODB.Recordset, rNa As New ADODB.Recordset
 Dim op$
 On Error GoTo fehler
' SET rFa = TabÖff("faelle", "Auswahl")
 myFrag rFa, "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_ID & " AND bhfe1 >= " & DatFor_k(Datu) & IIf(nurKasse, " AND schgr <> '90' ", "") & " ORDER BY bhfe1"
 If rFa.BOF Then ' wenn kein Fall in die Zeitschiene paßt, dann den letzten nehmen
  Set rFa = Nothing
  myFrag rFa, "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_ID & IIf(nurKasse, " AND schgr <> '90' ", "") & " ORDER BY bhfe1 DESC"
 End If
 If rFa.BOF Then Exit Function
 If Arztnr <> 0 And rFa!lanrid <> Arztnr Then Exit Function
 myFrag rNa, "SELECT * FROM `namen` WHERE pat_id = " & Pat_ID
 Lese.Ausgeb "Trage Leistung " & Leist & " für Pat_ID " & Pat_ID & " (" & rNa!Nachname & ", " & rNa!Vorname & ") mit Datum " & Format$(Datu, "dd.mm.yyyy") & " ein.", True
 LeistungsExport1a = 1
 Call BDT.Satzart(IIf(rFa!SchGr = "90", "0190", "0102")) ' 80000 Satzidentifikation
' bei 0101 entstehen bei zwei Aufrufen fehlerfrei zwei neue Kassenfaelle, jeder mit der Leistung
' bei 190 entsteht ein neuer Privatfall, bei 6100 läuft alles ohne Fehler durch, aber keine Leistung steht drin
' bei 6200 entsteht ein neuer Kassenfall
'     op = format$(3 + 4 + 4, "000") + "8000" + CStr(f!s8000)
'     Print #310, zsu(op)
'     op = format$(3 + 4 + 5, "000") + "8100" + CStr(f!s8100)
'     Print #310, zsu(op)
 Call BDT.PatID(rFa!Pat_ID)
'#If False THEN
'     IF 1 = 0 THEN 'Auswirkung bisher nicht geprüft 31.7.05 (3x)
'      op = Format$(3 + 4 + Len(n!NVorsatz), "000") + "3100" + CStr(n!NVorsatz)
'      Print #310, ZSU(op)
'     END IF
'     op = Format$(3 + 4 + Len(n!Nachname), "000") + "3101" + CStr(n!Nachname)
'     Print #310, ZSU(op)
'     op = Format$(3 + 4 + Len(n!Vorname), "000") + "3102" + CStr(n!Vorname)
'     Print #310, ZSU(op)
'     op = Format$(3 + 4 + 8, "000") + "3103" + Format$(n!GebDat, "ddmmyyyy")
'     Print #310, ZSU(op)
'     op = Format$(3 + 4 + Len(IIf(ISNULL(n!Versichertennummer), vNS, n!Versichertennummer)), "000") + "3105" + IIf(ISNULL(n!Versichertennummer), vNS, n!Versichertennummer)
'     Print #310, ZSU(op)
''     op = format$(3 + 4 + Len(n!Straße), "000") + "3107" + n!Straße
''     Print #310, ZSU(op)
''     op = format$(3 + 4 + Len(n!Plz), "000") + "3112" + n!Plz
''     Print #310, ZSU(op)
''     op = format$(3 + 4 + Len(n!Ort), "000") + "3113" + n!Ort
''     Print #310, ZSU(op)
'''     GoTo leistungdirekt
'     IF NOT ISNULL(n!KVKStatus) AND n!KVKStatus > 0 AND Trim$(n!KVKStatus) <> "" THEN
'      op = Format$(3 + 4 + Len(n!KVKStatus), "000") + "3108" + n!KVKStatus
'      Print #310, ZSU(op)
'     END IF
'     op = Format$(3 + 4 + 1, "000") + "3110" + IIf(n!Geschlecht = "m", "1", IIf(n!Geschlecht = "w", "2", ""))
'     Print #310, ZSU(op)
'#END IF
'#If False THEN
'     IF q!SchGr <> "90" THEN
'      op = Format$(3 + 4 + 5, "000") + "4101" + Quartal
'      Print #310, ZSU(op)
'      op = Format$(3 + 4 + 8, "000") + "4102" + Format$(rFa!ausgst, "ddmmyyyy")
'      Print #310, ZSU(op)
'      op = Format$(3 + 4 + Len(rFa!VKNr), "000") + "4104" + rFa!VKNr
'      Print #310, ZSU(op)
'     END IF
'#END IF
     BDT.SAdd "3635", "TM#" & rFa!Lanr
     BDT.SAdd "3636", "TM#" & BSNR
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
     End If ' rFa!SchGr <> "90" THEN
     If Not IsNull(rFa!GebOr) And rFa!GebOr <> vNS Then ' bei Privaten
      BDT.SAdd "4121", rFa!GebOr
     End If
     If Not IsNull(rFa!AbrGb) And rFa!AbrGb <> vNS Then ' bei Privaten
      BDT.SAdd "4122", rFa!AbrGb
     End If ' NOT ISNULL(rFa!GebOr) THEN ' bei Privaten
     BDT.SAdd "4144", "TM#" & IIf(IsNull(rFa!TMFNr), Space$(11), rFa!TMFNr)
     BDT.DAdd "4150", rFa!BhFB
     BDT.DAdd "4151", rFa!BhFE1
'     op = format$(3 + 4 + 8, "000") + "4152" + IIf(rFa!BhFE2 = 0, "00000000", format$(rFa!BhFE2, "ddmmyyyy"))
'     Print #310, ZSU(op)
     Dim Üw$
     BDT.SAdd "4218", "TM" & rFa!ÜbWVLANR & "#" & rFa!ÜbWVBSNR & "#" & rFa!ÜbWVKVNR
     Üw = rFa!Übwr ' IIf(ISNULL(rFa!AndÜw), vNS, rFa!AndÜw)
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
     BDT.SAdd "3636", "TM#" & BSNR
     BDT.SAdd "5001", Leist
 Call BDT.Schreib
 Close #310
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LeistungsExport1a/" + AnwPfad)
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
' IF DBCn.State = 0 THEN
'  Call DBVerb.cnVorb("", "anamnesebogen", "Patientendaten")
'  SetDBCn DBVerb.wCn
'  LVobMySQL = InStrB(UCase$(DBVerb.CnStr), "MYSQL") > 0
' END IF
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
  If InStrB(Datei, " PID ") <> 0 And InStrB(LCase$(Datei), ".pdf") <> 0 Then
   Dim geeignet%
   geeignet = geeignet + 1
   DateiDatum = FileDateTime(Verz & Datei)
   If DateiDatum >= fctQAnf(ZQuart(Now() - Verspätung)) And DateiDatum <= fctQEnd(ZQuart(Now() - Verspätung)) Then
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
     sql = "SELECT * FROM `leistungen` WHERE pat_id = " & pid & " AND leistung IN (40110,40111) AND " & SelDatum("zeitpunkt", Datum) ' "date(zeitpunkt) = " & DatFor_k(datum)
     Set rs = Nothing
     myFrag rs, sql
     If rs.BOF Then
      Zahl = Zahl + LeistungsExport1(BDT, pid, "40111", Datum, CDate("18:00"), True, Arztnr)
     End If
    ElseIf InStrB(Datei, "Arztbrief vom") <> 0 Then
     Pidpos = InStr(Datei, "Arztbrief vom ") + 14
     pidp2 = InStr(Pidpos, Datei, " ")
     Datum = Mid$(Datei, Pidpos, pidp2 - Pidpos)
'    Call LeistungsExport1(pid, "01601", datum, CDate("18:00"))
     sql = "SELECT * FROM leistungen WHERE pat_id = " & pid & " AND leistung IN (40110,40111) AND DATE(zeitpunkt) = " & DatFor_k(Datum)
     Set rs = Nothing
     myFrag rs, sql
     If rs.BOF Then
      Zahl = Zahl + LeistungsExport1(BDT, pid, "40110", Datum, CDate("18:00"), True, Arztnr)
     End If
    End If ' InStrB(Datei, " DMP-Daten ") <> 0 THEN elseif
   End If ' DateiDatum >= fctQAnf(ZQuart(NOW() - Verspätung)) AND DateiDatum <= fctQEnd(ZQuart(NOW() - Verspätung)) THEN
  End If ' InStrB(Datei, " PID ") <> 0 AND InStrB(LCase$(Datei), ".pdf") <> 0 THEN
  Datei = Dir
 Loop ' While LenB(Datei) <> 0
 
' sql = "SELECT b.* FROM (SELECT pat_id, DATE(zeitpunkt) AS datum, time(zeitpunkt) AS zeit FROM `tmbrie` b WHERE (name LIKE '%brief%' OR name LIKE '%nachricht%') AND zeitpunkt BETWEEN " & lQAnfuEnd(Str(Verspätung)) & " GROUP BY pat_id, DATE(zeitpunkt)) AS b LEFT JOIN `leistungen` l ON b.pat_id = l.pat_id  AND leistung LIKE '4012%' WHERE ISNULL(leistung)" ' DatFor_k(fctQAnf(ZQuart(NOW() - Verspätung))) & " AND " & DatFor_k(fctQEnd(ZQuart(NOW() - Verspätung)))
' ktag fehlerhaft
 sql = "SELECT b.* FROM (SELECT pat_id, DATE(zeitpunkt) Datum, TIME(zeitpunkt) zeit FROM `tmbrie` b WHERE (name LIKE '%brief%' OR name LIKE '%nachricht%') AND zeitpunkt BETWEEN  CONCAT(YEAR(NOW()-INTERVAL  25 DAY),'-',(QUARTER(NOW()-INTERVAL  25 DAY)-1)*3+1,'-01') AND CONCAT(YEAR(NOW()-INTERVAL  25 DAY)+ quarter(NOW()-INTERVAL  25 DAY) div 4 ,'-',((QUARTER(NOW()-INTERVAL  25 DAY)-1)*3+4) mod 12,'-01')-INTERVAL 0 day  GROUP BY pat_id, DATE(zeitpunkt)) AS b LEFT JOIN `leistungen` l ON b.pat_id = l.pat_id  AND DATE(l.zeitpunkt) = b.datum AND leistung LIKE '4012%' WHERE ISNULL(leistung)"
 Set rs = Nothing
 myFrag rs, sql
 Zahl1 = 0
 Do While Not rs.EOF
  Zahl1 = Zahl1 + LeistungsExport1(BDT, rs!Pat_ID, "40111", rs!Datum, CDate(rs!Zeit), True, Arztnr)
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
 Lese.Ausgeb "Datei '" & BDT.DMPImp & "' neu mit " & vbCrLf & Zahl & " Leistungen zu den Briefen in '" & Verz & "' und " & vbCrLf & Zahl1 & " Leistungen zu Arztbriefen in der Tabelle `tmbrie` für Kassenpatienten erstellt!, geeignet: " & geeignet, True
 Lese.Ausgeb sql, True
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doPorto/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dodoPorto

Function tuBriefeLeiDok(frm As Lese, Optional Arztnr&)
 On Error GoTo fehler
 Dim sql2$
 Dim rf As New ADODB.Recordset, rB As New ADODB.Recordset, rl As New ADODB.Recordset
 sql1 = "SELECT MIN(bhfb) AS bfb, MIN(schgr) AS schgr, pat_id, MIN(fid) AS fid FROM `faelle` WHERE schgr IN ('21','23','24','00') AND quartal = """ & ZQuart(Now - Verspätung) & """ GROUP BY pat_id"
' sql1 = "SELECT f.fid AS fid, DATE(b.zeitpunkt) AS tag, b.* FROM `tmbrie` b LEFT JOIN `faelle` f ON b.pat_id = f.pat_id WHERE quartal = '" & ZQuart(Now - Verspätung) & "' AND schgr IN ('21','23','24','00') AND b.zeitpunkt >= " & DatFor_k(fctQAnf(ZQuart(NOW() - 20))) & " AND name LIKE '%.doc' AND (name LIKE '%brief%' OR name LIKE '%dmp-daten%' OR name LIKE '%nachricht an%') ORDER BY b.pat_id,b.zeitpunkt;"
' Call LeistungsExport0
 myFrag rf, sql1
 Do While Not rf.EOF
'  IF rF!Pat_id = 2155 THEN
  sql1 = "SELECT * FROM `tmbrie` WHERE pat_id = " & rf!Pat_ID & " AND zeitpunkt >= " & DatFor_k(fctQAnf(ZQuart(Now() - 20))) & " AND name LIKE '%.doc' AND (name LIKE '%brief%' OR name LIKE '%nachricht an%') AND zeitpunkt >= " & DatFor_k(#10/16/2007#)
  Set rB = Nothing
'  rB.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
  myFrag rB, sql1
  If Not rB.EOF Then
   Set rl = Nothing
   myFrag rl, "SELECT * FROM `leistungen` WHERE fid = " & rf!FID & " AND leistung IN ('01601')"
   If rl.BOF Then
'    Call LeistungsExport1(rF!Pat_id, "01601", MIN(DateValue(rB!Zeitpunkt), #12/31/2007#), CDate("18:00"))
    Call LeistungsExport1(rf!Pat_ID, "40110", MINvb(DateValue(rB!Zeitpunkt), #12/31/2007#), CDate("18:00"), , , Arztnr)
   End If
  Else
   sql1 = "SELECT * FROM `tmbrie` WHERE pat_id = " & rf!Pat_ID & " AND zeitpunkt >= " & DatFor_k(fctQAnf(ZQuart(Now() - 20))) & " AND name LIKE '%.doc' AND (name LIKE '%dmp-daten%') AND zeitpunkt >= " & DatFor_k(#10/16/2007#)
   Set rB = Nothing
   myFrag rB, sql1
   If Not rB.EOF Then
    Set rl = Nothing
    myFrag rl, "SELECT * FROM `leistungen` WHERE fid = " & rf!FID & " AND leistung IN ('01600')"
    If rl.BOF Then
'     Call LeistungsExport1(rF!Pat_id, "01600", MIN(DateValue(rB!Zeitpunkt), #12/31/2007#), CDate("18:00"))
     Call LeistungsExport1(rf!Pat_ID, "40110", MINvb(DateValue(rB!Zeitpunkt), #12/31/2007#), CDate("18:00"), , , Arztnr)
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in tuBriefeLeiDok/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' tuBriefeLeiDok

Function alleDMPLeiDok(frm As Lese, Optional Arztnr&)
 Dim DefDB$
 Dim rsa As New ADODB.Recordset, rMV As New ADODB.Recordset
 On Error GoTo fehler
' sql1 = "SELECT schgr, `faelle`.bhfb AS bfb, `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.vorname AS hvv, `hae`.nachname AS hvn, `hae`.fax1, hausaerzte.telefax, üwnnr FROM (((`faelle` INNER JOIN `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(MID(`faelle`.üwnnr,1,2)¡""/""¡MID(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & DefDB & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr WHERE quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr IN (""24"",""23"",""21"",""00"") AND tkz = 0 ORDER BY nachname, vorname, schgr"
 sql1 = "SELECT bhfb AS bfb, schgr, pat_id FROM `faelle` WHERE schgr IN (""24"",""23"",""21"",""00"") AND quartal = """ & ZQuart(Now - Verspätung) & """ "
' IF lies.obMySQL THEN
'  sql1 = replace$(sql1, "¡", ",")
' Else
'  sql1 = replace$(replace$(sql1, "concat", ""), "¡", " & ")
' END IF
 myFrag rsa, "SELECT COUNT(DISTINCT pat_id) FROM (" & sql1 & ") innen"
 frm.Ausgeb rsa.Fields(0) & " Fälle ", True
 Set rsa = Nothing
' rsa.Open sql1, DBCn, adOpenDynamic, adLockOptimistic
 myFrag rsa, sql1
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
    OkSql = "SELECT * FROM ((((`forminhkopf` LEFT JOIN `formulare` ON `forminhkopf`.form_id = `formulare`.formid) LEFT JOIN `forminhfeld` ON `forminhfeld`.foid = `forminhkopf`.foid) LEFT JOIN `forminhaltfeld` ON `forminhaltfeld`.feldvw = `forminhfeld`.feldvw) LEFT JOIN `forminhaltfeldinh` ON `forminhaltfeldinh`.feldinhvw = `forminhfeld`.feldinhvw) LEFT JOIN `namen` ON `forminhkopf`.pat_id = `namen`.pat_id WHERE form_abk LIKE ""DMPDTYP" & "%" & """ AND feld LIKE """ & "%" & "datum"" AND `namen`.Pat_id = " & rsa!Pat_ID & ";"
    Set rMV = Nothing
'    rMV.Open OkSql, DBCn, adOpenDynamic, adLockReadOnly
    myFrag rMV, OkSql
    If rMV.EOF Then obzutr = 0
  End If
  If obzutr Then
    Set rMV = Nothing
'    rMV.Open "SELECT * FROM `leistungen` WHERE pat_id = " & rsa!Pat_id & " AND leistung = 1601 AND zeitpunkt >= " & DatFor_k(fctQAnf(ZQuart(Now - Verspätung))) & " AND zeitpunkt < " & DatFor_k(fctQEnd(ZQuart(Now - Verspätung))), DBCn, adOpenDynamic, adLockReadOnly
    myFrag rMV, "SELECT * FROM `leistungen` WHERE pat_id = " & rsa!Pat_ID & " AND leistung = 1601 AND zeitpunkt >= " & DatFor_k(fctQAnf(ZQuart(Now - Verspätung))) & " AND zeitpunkt < " & DatFor_k(fctQEnd(ZQuart(Now - Verspätung)))
    If Not rMV.BOF Then obzutr = 0
  End If
  If obzutr Then
   If anfang Then
    If rsa!Pat_ID <> AltID Then
' in MVZ nur die schicken, die nicht schon bei mir eingeschrieben
      Dim DaT As Date
      DaT = CDate("30.6.07")
      If rsa!bfb > DaT Then
       DaT = CDate("30.6.07")
      End If
'      Call LeistungsExport1(BDT, rsA!Pat_id, "01601", DaT, CDate("18:00"))
      Call LeistungsExport1(BDT, rsa!Pat_ID, "40110", DaT, CDate("18:00"), , Arztnr)
      FZahl = FZahl + 1
    End If
    AltID = rsa!Pat_ID
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in alleDMPLeiDok/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' alleDMPLeiDok

'Function alleDMPLeiDok1(frm AS Lese)
' Dim rsA As New ADODB.Recordset, rMV As New ADODB.Recordset
'
' sql1 = "SELECT `faelle`.bhfb AS bfb, `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.vorname AS hvv, `hae`.nachname AS hvn, `hae`.fax1, `hausaerzte`.telefax, üwnnr FROM (((`faelle` INNER JOIN `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(MID(`faelle`.üwnnr,1,2)¡""/""¡MID(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN `" & DefDB(DBCn) & "`.`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr WHERE quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr IN (""24"",""23"",""21"") AND tkz = 0 AND ((icd LIKE ""E11." & "%" & """ AND `hae`.DMPT2 <> 0) OR (icd LIKE ""E10." & "%" & """ AND `hae`.DMPT1 <> 0)) AND NOT üwnnr IN (""6493842"",""6496648"",""" & kvnr & """,""6491291"",""6488286"",""6420029"") ORDER BY nachname, vorname, schgr;"
''' Dr. colberg / Schmidt:
'' sql1 = "SELECT DISTINCT `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.vorname AS hvv, `hae`.nachname AS hvn, `hae`.fax1, `hausaerzte`.telefax, üwnnr FROM (((`faelle` INNER JOIN `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(MID(`faelle`.üwnnr,1,2)¡""/""¡MID(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & defdb & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr WHERE quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr IN (""24"",""23"",""21"") AND tkz = 0 AND ((icd LIKE ""E11." & "%" & """ AND `hae`.DMPT2 <> 0) OR (icd LIKE ""E10." & "%" & """ AND `hae`.DMPT1 <> 0)) AND üwnnr IN (""6419027"", ""6419421"",""6419418"") ORDER BY nachname, vorname, schgr;"
' ' Dr. Hofner 64/80054
'' sql1 = "SELECT DISTINCT `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.vorname AS hvv, `hae`.nachname AS hvn, `hae`.fax1, `hausaerzte`.telefax, üwnnr FROM (((`faelle` INNER JOIN `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(MID(`faelle`.üwnnr,1,2)¡""/""¡MID(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & defdb & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr WHERE quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr IN (""24"",""23"",""21"") AND tkz = 0 AND ((icd LIKE ""E11." & "%" & """ AND `hae`.DMPT2 <> 0) OR (icd LIKE ""E10." & "%" & """ AND `hae`.DMPT1 <> 0)) AND üwnnr IN (""6480054"") ORDER BY nachname, vorname, schgr;"
' IF lies.obMySQL THEN
'  sql1 = replace$(sql1, "¡", ",")
' Else
'  sql1 = replace$(replace$(sql1, "concat", ""), "¡", " & ")
' END IF
' rsA.Open sql1, DBCn, adOpenDynamic, adLockOptimistic
' Dim FZahl&, NZahl&
' Dim AltID&
' Dim anfang%
' Dim BDT As New BDTSchreib
' IF Not BDT.Start(hVerz, "Leist") THEN
'  Exit Function
' END IF
' Call BDT.ImportFolderHerricht
' Call BDT.BDTKopf
'' Call LeistungsExport0
' Do While Not rsA.EOF
'  IF anfang THEN
' ' IF (ISNULL(rsA!telefax) AND ISNULL(rsA!fax1)) OR (NOT ISNULL(rsA!telefax) AND NOT ISNULL(rsA!fax1) AND rsA!telefax <> rsA!fax1) THEN
'   IF rsA!Pat_id <> AltID THEN
'' in MVZ nur die schicken, die nicht schon bei mir eingeschrieben
'    Dim OkSql$
'    OkSql = "SELECT * FROM ((((`forminhkopf` LEFT JOIN `formulare` ON `forminhkopf`.form_id = `formulare`.formid) LEFT JOIN `forminhfeld` ON `forminhfeld`.foid = `forminhkopf`.foid) LEFT JOIN `forminhaltfeld` ON `forminhaltfeld`.feldvw = `forminhfeld`.feldvw) LEFT JOIN `forminhaltfeldinh` ON `forminhaltfeldinh`.feldinhvw = `forminhfeld`.feldinhvw) LEFT JOIN `namen` ON `forminhkopf`.pat_id = `namen`.pat_id WHERE form_abk LIKE ""DMPDTYP" & "%" & """ AND feld LIKE """ & "%" & "datum"" AND `namen`.Pat_id = " & rsA!Pat_id & ";"
'    SET rMV = Nothing
'    rMV.Open OkSql, DBCn, adOpenDynamic, adLockReadOnly
'    IF rMV.EOF OR rsA!ÜWNNr <> "6419416" THEN ' beim MVZ nicht die schon bei uns eingetragenen
'     SET rMV = Nothing
'     rMV.Open "SELECT * FROM `leistungen` WHERE leistung = 1600 AND zeitpunkt >= " & DatFor_k(CDate("28.6.07")) & " AND zeitpunkt < " & DatFor_k(CDate("28.6.07") + 1) & " AND pat_id = " & rsA!Pat_id, DBCn, adOpenDynamic, adLockReadOnly
'     IF rMV.BOF THEN
'      Dim DaT As Date
'      DaT = CDate("16.6.07")
'      IF rsA!bfb > DaT THEN
'       DaT = CDate("30.6.07")
'      END IF
''      Call LeistungsExport1(BDT, rsA!Pat_id&, "01600", DaT, CDate("13:00"))
'      Call LeistungsExport1(BDT, rsA!Pat_id&, "40120", DaT, CDate("13:00"))
'      FZahl = FZahl + 1
'     Else
'      NZahl = NZahl + 1
'     END IF
'    END IF
'    AltID = rsA!Pat_id
'   END IF
'  END IF
''  IF rsA!Pat_id = 977 THEN
'   anfang = True
''  endif
'  rsA.Move 1
' Loop
' frm.Ausgeb "In " & FZahl & " Fällen Leistungen eingetragen, in " & NZahl & " weggelassen.", True
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in alleDMPLeiDok1/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' alleDMPLeiDok1

' in GetMed, dolies, rezEintr, do_Pat_ID_Change, PatAuswahl, Command1_Click, callMachDMPBogen, Form_Load (Pat_Liste), MFG_Mouse_Down, UKPDS, dodoPLZ, LaborIns1, obhierdmpfn, TabAusgeb
Function MAXvb(D1, D2)
 If D1 > D2 Then MAXvb = D1 Else MAXvb = D2
End Function ' max

' in liesExcel, std3, Command1_Click, doFS, callMachDMPBogen, dodoPLZ, DMPString, tuBriefeLeiDok, einDMP, getPrRR, TabAusgeb, Form_Load (LANRAuswahl),
Function MINvb(D1, D2)
 If D1 < D2 Then MINvb = D1 Else MINvb = D2
End Function ' min

#If zutesten Then
' kommt nirgends vor
Sub tobhier()
 Dim dmpk&, dmpkhkk&, dmpcopdk&, dmpabk&, HzV&, Ds&, hzvab As Date, dsab As Date
 Dim dmpb As Date, dmpkhkb As Date, dmpcopdb As Date, dmpabb As Date, NZNr&
 Call obhierdmpfn("DMP COPD MVZ 12.10.19", NZNr, dmpk, dmpb, dmpkhkk, dmpkhkb, dmpcopdk, dmpcopdb, dmpabk, dmpabb, HzV, hzvab, Ds, dsab)
' Debug.Print dmpk, dmpb, dmpcopdk, dmpcopdb, HzV, hzvab, DS, dsab
End Sub ' tobhier
#End If

' in alleSpeichern() und tob()
Sub obhierdmpfn(notiz$, NZNr&, dmpklass As DMPEnum, dmpbeg As Date, Optional dmpkhkklass&, Optional dmpkhkbeg As Date, Optional dmpcopdklass&, Optional dmpcopdbeg As Date, Optional dmpabklass&, Optional dmpabbeg As Date, Optional HzV&, Optional HzVbeg As Date, Optional Ds&, Optional DSbeg As Date)
 Dim buch$, pos%, testd As Date, i%, notdat$(), ndRichtig%() ' letztes wahr, wenn notdat-Zeile für DMP verwertbar oder irrelevant
 Dim maxdHier(3) As Date, maxdHA(3) As Date, maxdNein(3) As Date, maxdAus(3) As Date
 Dim DMPArt% ' 0 = Diabetes, 1 = KHK, 2 = COPD, 3 = Asthma bronchiale
 On Error GoTo fehler
    If LenB(Trim$(notiz)) <> 0 Then
     ReDim notdat(0)
     Do
      pos = InStr(notiz, vbCrLf)
      If pos <> 0 Then
       notdat(UBound(notdat)) = UCase$(Trim$(left$(notiz, pos - 1)))
       notiz = Mid$(notiz, pos + 2)
       ReDim Preserve notdat(UBound(notdat) + 1)
      Else ' pos <> 0 Then
       notdat(UBound(notdat)) = UCase$(Trim$(notiz))
       notiz = vNS
       Exit Do
      End If ' pos <> 0 Then else
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
       Ds = 1
       DSbeg = testd
'      IF InStrB(notdat(i), "DMP") = 0 THEN
'       ndRichtig(i) = True ' a) für DMP irrelevant
      ElseIf InStrB(notdat(i), "DMP") = 1 Then
       DMPArt = 0
       If InStrB(notdat(i), "KHK") <> 0 Then DMPArt = 1 Else If InStrB(notdat(i), "COPD") <> 0 Then DMPArt = 2 Else If (InStrB(notdat(i), "ASTHMA") <> 0 Or InStrB(notdat(i), "AB") <> 0) Then DMPArt = 3
       If testd <> 0 Then
        If InStrB(notdat(i), "HIER") <> 0 Or InStrB(notdat(i), "BEI UNS") <> 0 Then
         maxdHier(DMPArt) = MAXvb(maxdHier(DMPArt), testd)
        ElseIf InStrB(notdat(i), " HA") <> 0 Or InStrB(notdat(i), " HÄ") <> 0 Then
         maxdHA(DMPArt) = MAXvb(maxdHA(DMPArt), testd)
        ElseIf InStrB(notdat(i), "NEIN") <> 0 Then
         maxdNein(DMPArt) = MAXvb(maxdNein(DMPArt), testd)
        ElseIf InStrB(notdat(i), "AUSGESCHRIEBEN") <> 0 Then
         maxdAus(DMPArt) = MAXvb(maxdAus(DMPArt), testd)
        Else ' z.B. DMP MVZ
         maxdHA(DMPArt) = MAXvb(maxdHA(DMPArt), testd)
         ndRichtig(i) = False ' b2) und "hier", " HA" oder "nein" enthalten
        End If ' instrb(notdat(i), ...
       End If ' test<>0
       If InStrB(notdat(i), "NEIN") <> 0 Then
        maxdNein(DMPArt) = Now() - 365
       ElseIf InStrB(notdat(i), "AUSGESCHRIEBEN") <> 0 Then
        maxdAus(DMPArt) = Now() - 365
       End If ' InStrB(notdat(i), "NEIN") <> 0 THEN
      End If ' instrb(notdat(i), ...
     Next i
     For i = UBound(notdat) To 0 Step -1
      If ndRichtig(i) = 0 Then
       notiz = notdat(i) & vbCrLf & notiz
       NZNr = i
      End If ' ndRichtig(i) = 0 Then
     Next i
     For DMPArt = 3 To 0 Step -1
        dmpklass = unb
        If maxdAus(DMPArt) > maxdHier(DMPArt) And maxdAus(DMPArt) > maxdHA(DMPArt) And maxdAus(DMPArt) > maxdHA(DMPArt) And maxdAus(DMPArt) <> 0 Then
         dmpklass = ausg
         dmpbeg = maxdAus(DMPArt)
        ElseIf maxdHier(DMPArt) > maxdHA(DMPArt) And maxdHier(DMPArt) >= maxdNein(DMPArt) And maxdHier(DMPArt) <> 0 Then
          dmpklass = hier
          dmpbeg = maxdHier(DMPArt)
        ElseIf maxdNein(DMPArt) > maxdHA(DMPArt) And maxdNein(DMPArt) <> 0 Then
          dmpklass = nein
          dmpbeg = maxdNein(DMPArt)
        ElseIf maxdHA(DMPArt) <> 0 Then
          dmpklass = HA
          dmpbeg = maxdHA(DMPArt)
        End If ' maxdAus( ...
        If dmpklass <> unbek Then
         If DMPArt = 1 Then
          dmpkhkklass = dmpklass
          dmpkhkbeg = dmpbeg
          dmpklass = unb
          dmpbeg = 0
         ElseIf DMPArt = 2 Then
          dmpcopdklass = dmpklass
          dmpcopdbeg = dmpbeg
          dmpklass = unb
          dmpbeg = 0
         ElseIf DMPArt = 3 Then
          dmpabklass = dmpklass
          dmpabbeg = dmpbeg
          dmpklass = unb
          dmpbeg = 0
         End If
        End If
     Next DMPArt
    End If ' LenB(Notiz) <> 0 THEN
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in obhierdmpfn/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub      ' obhierdmpfn

Function doVerdächtigeÜberweiser()
 Dim sql$, rf As New ADODB.Recordset, infos$(), taskid&, HAi%, aktQ$, obverd%
 Dim Datei$
 Datei$ = pVerz & "VerdächtigeÜberweiser.txt"
 aktQ = left$(ZQuart(Now() - 14), 1) & Right$(ZQuart(Now() - 14), 2) ' 109
 sql = "SELECT v.pat_id,n.* FROM `aktfvs` v LEFT JOIN `namen` n ON v.pat_id = n.pat_id"
 Open Datei For Output As #355
' rf.Open sql, DBCn
 myFrag rf, sql
 Do While Not rf.EOF
'   Call getHausarztAlt(Pid:=rF!Pat_id, Infos:=Infos(), obHAPrio:=False)
   Dim rFa() As Faelle
   Dim rKv1() As kvnrue
   Call getHausarzt1(infos, rFa, rKv1, False, rf!Pat_ID, , , "doVerdächtigeÜberweiser")
   HAi = 0
   obverd = False
   Do
    If InStrB(infos(10, HAi), "HA") <> 0 Then Exit Do
    HAi = HAi + 1
    If HAi > UBound(infos, 2) Then obverd = True: Exit Do
   Loop
   If Not obverd Then If HAi <> 0 And InStrB(infos(10, 0), aktQ) <> 0 And infos(1, 0) <> infos(1, HAi) Then obverd = True
   If obverd Then
    Print #355, rf!Pat_ID & ": " & rf!Nachname & ", " & rf!Vorname & ":"
    Print #355, "  " & left$(infos(10, 0) & Space$(10), 10) & ": " & infos(1, 0)
    If Not HAi > UBound(infos, 2) Then Print #355, "  " & left$(infos(10, HAi) & Space$(10), 10) & ": " & infos(1, HAi)
    Ausgeb "verdächtig: " & rf!Pat_ID & ": " & rf!Nachname & ", " & rf!Vorname, -1
   Else
    Ausgeb "in Ordnung: " & rf!Pat_ID & ": " & rf!Nachname & ", " & rf!Vorname, -1
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
' Dim hae As New adodb.Recordset
' Dim rsA As New adodb.Recordset, rMV As New adodb.Recordset
' Dim sql$, sql1$
' Dim AltID&
' Dim antw&
' antw = MsgBox("Wollen Sie wirklich alle DMP-Daten versenden?", vbYesNo)
' IF antw = vbNo THEN Exit Function
' QMdbAkt = QmdB ' Imortiere.bas uverz & "anamnese\quelle.mdb"
''SELECT `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.telefax AS hfax, `hausaerzte`.telefax FROM (((`faelle` INNER JOIN `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(MID(`faelle`.üwnnr,1,2)¡"/"¡MID(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & defdb & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr WHERE quartal = "42006" AND schgr IN ("24","23","21") AND icd LIKE "E11.%" AND tkz = 0 AND `hae`.DMPT2 <> 0 AND NOT üwnnr IN ("6493842","6496648","" & kvnr & "","6491291","6488286","6420029","6494531") ORDER BY schgr;
''Ausnahmen: Lindenthal, Pfeuffer, Rembold, Schneider, Stolzki, Wengeler
'
''  LEFT JOIN `kvaerzte`.`hae` ON CONCAT(mid$(`faelle`.üwnnr,1,2)¡""/""¡mid$(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & DefDB & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr) LEFT JOIN " & DefDB & ".`liuez` AS haxls ON `faelle`.üwnnr = haxls.kvnr
'
' sql1 = "SELECT `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, üwnnr, `namen`.info, `namen`.obhierdmp, icd FROM (((`" & DefDB(DBCn) & "`.`faelle` INNER JOIN `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `namen` ON `faelle`.pat_id = `namen`.pat_id) WHERE (quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr IN (""00"",""24"") AND tkz = 0 AND icd REGEXP '^E1[0-4]\.' AND obhierdmp=0) ORDER BY nachname, vorname, schgr"
''' Dr. colberg / Schmidt:
''  sql1 = "SELECT DISTINCT `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.vorname AS hvv, `hae`.nachname AS hvn, `hae`.fax1, `hausaerzte`.telefax, haxls.fax, üwnnr, `namen`.info FROM (((((`faelle` INNER JOIN `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(MID(`faelle`.üwnnr,1,2)¡""/""¡MID(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & defdb & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr) LEFT JOIN " & defdb & ".`liuez` AS haxls ON `faelle`.üwnnr = haxls.kvnr) LEFT JOIN `namen` ON `faelle`.pat_id = `namen`.pat_id WHERE quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr IN (""24"",""23"",""21"") AND tkz = 0 AND ((icd LIKE ""E11." & "%" & """ AND `hae`.DMPT2 <> 0) OR (icd LIKE ""E10." & "%" & """ AND `hae`.DMPT1 <> 0)) AND üwnnr IN (""6419027"", ""6419421"",""6419418"") ORDER BY nachname, vorname, schgr"
' ' Dr. Hofner 64/80054
'' sql1 = "SELECT DISTINCT `faelle`.Pat_id, `faelle`.nachname, `faelle`.vorname, `hae`.vorname AS hvv, `hae`.nachname AS hvn, `hae`.fax1, `hausaerzte`.telefax, haxls.fax, üwnnr, `namen`.info FROM (((((`faelle` INNER JOIN `diagnosen` ON `faelle`.pat_id = `diagnosen`.pat_id) LEFT JOIN `anamnesebogen` ON `faelle`.pat_id = `anamnesebogen`.pat_id) LEFT JOIN `kvaerzte`.`hae` ON CONCAT(MID(`faelle`.üwnnr,1,2)¡""/""¡MID(`faelle`.üwnnr,3,5)) = `hae`.KVNR) LEFT JOIN " & defdb & ".`hausaerzte` ON `faelle`.üwnnr = `hausaerzte`.kvnr) LEFT JOIN " & defdb & ".`liuez` AS haxls ON `faelle`.üwnnr = haxls.kvnr) LEFT JOIN `namen` ON `faelle`.pat_id = `namen`.pat_id WHERE quartal = """ & ZQuart(Now - Verspätung) & """ AND schgr IN (""24"",""23"",""21"") AND tkz = 0 AND ((icd LIKE ""E11." & "%" & """ AND `hae`.DMPT2 <> 0) OR (icd LIKE ""E10." & "%" & """ AND `hae`.DMPT1 <> 0)) AND üwnnr IN (""6480054"") ORDER BY nachname, vorname, schgr"
' IF lies.obMySQL THEN
'  sql1 = replace$(sql1, "¡", ",")
' Else
'  sql1 = replace$(replace$(sql1, "concat", ""), "¡", " & ")
' END IF
' myEFrag "DROP TABLE dmpausw"
' myEFrag "CREATE TABLE dmpausw(pat_id int(10) unique key,name varchar(100), üwnnr varchar(10), icd varchar(10),fax varchar(20), adressat varchar(150)) comment 'Auswahl für alleDMPs'"
' myEFrag "DELETE FROM dmpausw"
' sql1 = "SELECT * FROM (" & sql1 & ") AS innen GROUP BY pat_id" ' WHERE pat_id = 2193
' rsA.Open sql1, DBCn, adOpenDynamic, adLockOptimistic
'' For runde = 1 To 2
' Do While Not rsA.EOF
'   IF rsA!Pat_id <> AltID THEN
'    IF True OR rsA!Pat_id = 50 THEN
' '    IF rsA!obhierdmp <> 0 THEN ' falls nicht bei uns unterschrieben
'      lies.Ausgeb rsA!Pat_id, False, True
'      Call einDMP(rsA!Pat_id, rsA!ICD, dszahl, rsA!Nachname, rsA!Vorname, rsA!ÜWNNr, frm)
''     END IF
'    END IF
'   END IF
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
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in alleDMPs/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' allDMPs

' in Lese.los
Function einDMP(Pat_ID&, Optional ICD$, Optional dszahl&, Optional Nachname$, Optional Vorname$, Optional ÜWNNr$, Optional frm As Lese)
 Dim Faxnr$, infos$() ' Frau/Herrn, Vorn+Nachn, Straße, PLZ+Ort, Faxnr, S.g./Liebe, DMPTyp2, DMPTyp1
 Dim aktPatGefaxt$()
 Dim i%, j&, obdoppelt%, rAf&
 Dim docName$, Adressat$, fax1$, anfang%, runde%
 Const maxverschieden% = 1 ' maximale Zahl verschiedener Ärzte, an die Infos geschickt werden
    ReDim aktPatGefaxt(0)
     Call getHausarzt(pid:=Pat_ID, infos:=infos()) ' , obHAPrio:=True)
'     Dim rs As New adodb.Recordset
'     myFrag rs, "SELECT getha FROM `namen` WHERE pat_id = " & Pat_id
'     IF rs.EOF THEN Stop
'     IF rs!getHA <> Infos(12, 0) THEN Stop
'     Exit Function
     If LenB(ICD) = 0 Then ICD = IIf(infos(6, i) = "X", "E11", "E10")
     For i = 0 To MINvb(maxverschieden - 1, UBound(infos, 2))
'      IF ((ICD LIKE "E11*" AND infos(6, i) = "X") OR (ICD LIKE "E10*" AND infos(7, i) = "X")) THEN
      If ((ICD Like "E11*") Or (ICD Like "E10*")) Then
      ' 5.7.10: Diese beiden Zeilen wären notwendig für die zwischenzeitliche Umfunktionierung dieser Funktion,
      ' jetzt in Tabellen `namen` (obhierdmp, obdmp, getha0, getha1, getha2) und hareal
'       InsKorr DBCn, "INSERT INTO `dmpausw` VALUES(" & Pat_id & ",'" & UmwfSQL(Nachname & ", " & Vorname) & "','" & Infos(12, i) & "','" & ICD & "','" & Infos(4, i) & "','" & UmwfSQL(Infos(14, i) & ", " & Infos(9, i) & ", " & Infos(3, i)) & "')", rAF
'       Exit Function
       Select Case infos(12, i)
        Case "6492325" ' Reitmeier, will keine DMP-Auswertungen mehr (23.12.09)
        Case KVNr ' Schade
'        Case "6493950", "6419568" ' Schorten, Sewering
'        Case "6495818", "6491291" ' Pfeuffer, Krombholz
        Case Else
         fax1 = infos(4, i)
         obdoppelt = False
         For j = 0 To UBound(aktPatGefaxt) - 1
          If aktPatGefaxt(j) = fax1 Then obdoppelt = True
         Next j
         If Not obdoppelt Then
'          Debug.Print "Erstelle: ", Pat_id, Nachname, Vorname, fax1, ÜWNNr
'          IF runde = 1 THEN
'          Else
'           IF pat_id = 1974 THEN
           anfang = True
'           END IF
           If anfang Then
'            Call Ausgeb(Pat_id & ": " & Nachname & ", " & Vorname & ", " & fax1 & ", " & ÜWNNr)
            DoEvents
            Adressat = infos(10, i) & " " & infos(1, i)
            docName = do_DMPAusgebStandAlone(Pat_ID, fax1, Adressat)
            If False Then
             Call FaxSend(docName, Adressat, fax1)
            End If
           End If ' 1 = 0
           aktPatGefaxt(UBound(aktPatGefaxt)) = fax1
           ReDim Preserve aktPatGefaxt(UBound(aktPatGefaxt) + 1)
'         Exit For ' Alternative zu obdoppelt
'          END IF
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in einDMPs/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' einDMP

Function doAnwalt(Pat_ID&)
 Dim rBr As New ADODB.Recordset, nr&, Datei$
 On Error GoTo fehler
 nr = 1
 myFrag rBr, "SELECT Pfad,Name,Art,Zeitpunkt FROM tmbrie WHERE pat_id=" & Pat_ID & " ORDER BY ZEITPUNKT"
 Do While Not rBr.EOF
  Datei = zVerz & Mid$(rBr!Pfad, 3)
  If FSO.FileExists(Datei) Then
   FSO.CopyFile Datei, pVerz & "anwalt\Dokument_" & nr & "_" & rBr!name & IIf(rBr!art = "wbr", ".doc", IIf(rBr!art = "pdf", ".pdf", ""))
   Lese.Ausgeb "Dokument '" & rBr!name & "', vom " & rBr!Zeitpunkt & ", Pfad: " & rBr!Pfad & " kopiert!", False
   nr = nr + 1
  Else
   Lese.Ausgeb "Dokument '" & rBr!name & "', vom " & rBr!Zeitpunkt & ", Pfad: " & rBr!Pfad & " nicht gefunden!", True
  End If
  rBr.MoveNext
 Loop ' While Not rBr.EOF
 Lese.Ausgeb nr & " Dokumente kopiert", True
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doAnwalt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Anwalt

#If False Then
Function Ausgeb(AGStr$)
 Dim runde%, aktdat$
 aktdat = AutoBriefProtok
 On Error Resume Next
 For runde = 0 To 10
  Open aktdat For Append As #334 ' pverz & "zufaxen.txt" = AutoBriefProtok
  If Err.Number = 0 Then Exit For
  aktdat = left$(aktdat, Len(aktdat) - 4) & "_" & Right$(aktdat, 4)
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Ausgeb/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Ausgeb(AGStr$)
#End If

'' SpMinÜ und SpMaxÜ = Arrays mit minimalen / maximalen Überschriftenlängen
'Public FUNCTION TabAusgeb(rEinl As Adodb.Recordset, Optional obMitausgeb% = False, Optional nz$ = vbCrLf, Optional ohneKopfz% = False, Optional SpMinÜ, Optional SpMaxÜ, Optional mitLeerZeilen% = False, Optional Datei$) AS CString
' Dim i%, maxL%(), Zrm%(), notNum%(), F1alt
' Const ZrmVorgabe% = 2
' ON Error GoTo fehler
' SET TabAusgeb = New CString
' ReDim maxL(rEinl.Fields.Count)
' ReDim Zrm(rEinl.Fields.Count)
' ReDim notNum(rEinl.Fields.Count)
' For i = 0 To UBound(Zrm)
'  Zrm(i) = ZrmVorgabe
' Next i
' Do While Not rEinl.EOF
'  For i = 0 To rEinl.Fields.Count - 1
'   IF Len(rEinl.Fields(i).Value) > maxL(i) THEN maxL(i) = Len(rEinl.Fields(i).Value)
'   IF NOT ISNULL(rEinl.Fields(i).Value) THEN IF Not IsNumeric(rEinl.Fields(i).Value) THEN notNum(i) = True
'  Next i
'  rEinl.Move 1
' Loop
' IF ohneKopfz = 0 THEN
'  For i = 0 To rEinl.Fields.Count - 1
'   IF Len(rEinl.Fields(i).Name) > maxL(i) THEN maxL(i) = Len(rEinl.Fields(i).Name): Zrm(i) = ZrmVorgabe - 1
'  Next i
' END IF
' IF Not IsMissing(SpMinÜ) THEN
'  For i = 0 To rEinl.Fields.Count - 1
'   IF UBound(SpMinÜ) >= i THEN
'    IF maxL(i) < SpMinÜ(i) AND SpMinÜ(i) <> 0 THEN
'     maxL(i) = SpMinÜ(i)
'    END IF
'   END IF
'  Next i
' END IF
' IF Not IsMissing(SpMaxÜ) THEN
'  For i = 0 To rEinl.Fields.Count - 1
'   IF UBound(SpMaxÜ) >= i THEN
'    IF maxL(i) > SpMaxÜ(i) AND SpMaxÜ(i) <> 0 THEN
'     maxL(i) = SpMaxÜ(i)
'    END IF
'   END IF
'  Next i
' END IF
' TabAusgeb.Clear
' IF ohneKopfZ = 0 THEN
'  For i = 0 To rEinl.Fields.Count - 1
'   TabAusgeb.Append Left$(rEinl.Fields(i).Name & Space$(maxL(i) + Zrm(i)), maxL(i) + Zrm(i))
'  Next i
'  TabAusgeb.Append nz
' END IF ' not ohneKopfZ
' IF Not rEinl.BOF THEN
'  rEinl.MoveFirst
'  IF mitLeerZeilen THEN F1alt = rEinl.Fields(0)
'  Do While Not rEinl.EOF
'   IF mitLeerZeilen THEN IF rEinl.Fields(0) <> F1alt THEN TabAusgeb.Append nz: F1alt = rEinl.Fields(0)
'   For i = 0 To rEinl.Fields.Count - 1
'    IF notNum(i) THEN
'     TabAusgeb.Append Left$(rEinl.Fields(i).Value & Space$(maxL(i) + Zrm(i)), maxL(i) + Zrm(i))
'    Else
'     TabAusgeb.Append Right$(Space$(maxL(i)) & rEinl.Fields(i).Value, maxL(i)) & Space$(Zrm(i))
'    END IF
'   Next i
'   TabAusgeb.Append nz
'   rEinl.Move 1
'  Loop
'  rEinl.MoveFirst
' END IF ' not rEinl.BOF
' IF obMitausgeb THEN Lese.Ausgeb TabAusgeb.Value, True
' IF LenB(Datei) <> 0 THEN
'  IF InStrB(Datei, ".") = 0 THEN
'   Datei = pVerz & Datei & " " & Format$(Now, "dd.mm.yy hh.mm.ss") & ".csv"
'  END IF
'  Open Datei For Output AS #317
'  Print #317, TabAusgeb.Value
'  Close #317
'  zeigan Datei
' END IF
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + nz + "LastDLLError: " + CStr(Err.LastDllError) + nz + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + nz + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in TabAusgeb/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' TabAusgeb

' in doFollowUp_Click
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
' Call rP.Open("SELECT DISTINCT f.pat_id AS pat_id FROM `faelle` f LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id WHERE not schgr IN ('41','42','43')  AND tkz = 0 ORDER BY f.pat_id DESC;", DBCn, adOpenDynamic, adLockReadOnly)
 myFrag rP, "SELECT DISTINCT f.pat_id AS pat_id FROM `faelle` f LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id WHERE not schgr IN ('41','42','43')  AND tkz = 0 ORDER BY f.pat_id DESC;"
 Do While Not rP.EOF
  Call getHausarzt(rP!Pat_ID, infos())
  If infos(1, 0) <> vNS And Not infos(1, 0) Like "*Schade" Then
   Set rB = Nothing
   sql = "SELECT * FROM `tmbrie` WHERE pat_id = " & rP!Pat_ID & " AND ((name LIKE '%Brief an %Dr%' OR name LIKE '%Arztbrief%' OR name LIKE 'Brief an HA%' OR name LIKE 'Brief an HAe%' OR name LIKE 'Brief an %') AND name NOT LIKE '%Entwurf%') ORDER BY zeitpunkt DESC"
'   Call rB.Open(sql, DBCn, adOpenDynamic, adLockReadOnly)
   myFrag rB, sql
   If Not rB.BOF Then
    If DateValue(lebe(rP!Pat_ID)) - rB!Zeitpunkt > 365 Then
     Lese.Aktion = Briefschreiben
     Call tuBriefStandalone(rP!Pat_ID, True, , , , , , , , True)
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dodoFollowUp/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' dodoFollowUp(frm AS Lese)

' in UngeschriebeneBriefe_Click
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
' Call rP.Open("SELECT DISTINCT f.pat_id AS pat_id FROM `faelle` f LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id WHERE not schgr IN ('41','42','43')  AND tkz = 0  ORDER BY f.pat_id DESC;", DBCn, adOpenDynamic, adLockReadOnly)
 myFrag rP, "SELECT DISTINCT f.pat_id AS pat_id FROM `faelle` f LEFT JOIN `anamnesebogen` a ON f.pat_id = a.pat_id WHERE not schgr IN ('41','42','43')  AND tkz = 0  ORDER BY f.pat_id DESC;"
 Do While Not rP.EOF
  Call getHausarzt(rP!Pat_ID, infos())
  If infos(1, 0) <> vNS And Not infos(1, 0) Like "*Schade" Then
   Set rB = Nothing
   sql = "SELECT * FROM `tmbrie` WHERE pat_id = " & rP!Pat_ID & " AND ((name LIKE '%Brief an %Dr%' OR name LIKE '%Arztbrief%' OR name LIKE 'Brief an HA%' OR name LIKE 'Brief an HAe%' OR name LIKE 'Brief an %') AND name NOT LIKE '%Entwurf%')"
'   Call rB.Open(sql, DBCn, adOpenDynamic, adLockReadOnly)
   myFrag rB, sql
   If rB.BOF Then
    Lese.Aktion = Briefschreiben
    Call tuBriefStandalone(rP!Pat_ID, True, , , , , , , , True)
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doUngeschriebeneBriefe/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' UngeschriebeneBriefe_Click

' in los()
Sub doRestlicheBriefe(frm As Lese, ab&)
 Dim sql$, sql1$
 Dim rP As New ADODB.Recordset
 Dim rB As New ADODB.Recordset
 Dim Zp As Date, Quartal$
 On Error GoTo fehler
 Zp = Now - 30
 Quartal = ZQuart(Zp)
' ab = InputBox("Ab welchem Pat. anfangen (0 = von vorne)?", "Rückfrage", 0)
' Call rP.Open("SELECT DISTINCT pat_id FROM `faelle` WHERE quartal = '" & Quartal & "' AND pat_id >= " & ab & " ORDER BY pat_id ", DBCn, adOpenDynamic, adLockReadOnly)
 myFrag rP, "SELECT DISTINCT pat_id FROM `faelle` WHERE quartal = '" & Quartal & "' AND pat_id >= " & ab & " ORDER BY pat_id "
 Do While Not rP.EOF
  Set rB = Nothing
  sql = "SELECT * FROM `tmbrie` WHERE pat_id = " & rP!Pat_ID & " AND zeitpunkt > " & DatFor_k(fctQAnf(Quartal)) & " AND ((name LIKE '%Brief an %Dr%' OR name LIKE '%Arztbrief%' OR name LIKE 'Brief an HA%' OR name LIKE 'Brief an HAe%') AND name NOT LIKE '%Entwurf%')"
'  Call rB.Open(sql, DBCn, adOpenDynamic, adLockReadOnly)
  myFrag rB, sql
  If rB.BOF Then
   Lese.Aktion = Briefschreiben
   Call tuBriefStandalone(rP!Pat_ID, True, , , , , , , , True)
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doRestlicheBriefe/" + AnwPfad)
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FaxSend/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' FaxSend


' in einDMP, doStart
Public Function do_DMPAusgebStandAlone(Pat_ID&, Optional fax1$, Optional Adressat$)
 Dim docName$, DT As DMPClass
 On Error GoTo fehler
 If fax1 = "08131 25169" Then fax1 = "08131 273373" ' Dr. Stolzki falsche Faxnummer
 If InStrB(fax1, "6119199") <> 0 Or InStrB(fax1, "91191") Then fax1 = "6119198"
 FNr = 0
 Dim rsAnam As New ADODB.Recordset
 Dim sverz$, guv$, dxml$
#If Not wordalt Then
 Const uvz$ = "2145\", uvuv$ = uvz & "neu\"
 guv = pVerz & "\zufaxen\" & uvuv
 dxml$ = guv & "word\document.xml"
#Else
 Const uvz$ = "2145\"
#End If
' wenn kein Adressat, dann in P:\ speichern
 If fax1$ = vNS Then
  sverz = pVerz
 Else
  sverz = pVerz & "zufaxen\" & uvz
 End If ' fax1 = vns
 myFrag rsAnam, "SELECT CONCAT('" & UmwfSQL(sverz) & "',COALESCE(MAX(gesname(pat_id)),''),', PID ','" & Pat_ID & "','" & ", DMP-Daten vom " & Format$(Now, "DD/MM/YY hh.mm.ss") & IIf(Adressat <> "" And Not IsNull(Adressat), " für " & Adressat, "") & IIf(LenB(fax1) = 0, vNS, " an Fax " & REPLACE$(fax1, "\N", "")) & ".doc') FROM `anamnesebogen` WHERE pat_id = " & Pat_ID
 If Not rsAnam.BOF Then
  docName = rsAnam.Fields(0)
 End If
' docName = sverz + rsAnam!Nachname & " " & rsAnam!Vorname & ", PID " & rsAnam!Pat_id & ", DMP-Daten vom " & Format$(Now, "DD/MM/YY hh.mm.ss") & IIf(Adressat <> "" And Not IsNull(Adressat), " für " & Adressat, "") & IIf(LenB(fax1) = 0, vNS, " an Fax " & fax1) & ".doc"
' #Const wordalt = True
#If Not wordalt Then
 Const AbsE$ = "</w:t></w:r></w:p>"
 Const AbE1$ = "<w:p><w:pPr><w:pStyle w:val=""Normal""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/><w:tab w:val=""left"" w:pos=""3118"" w:leader=""dot""/></w:tabs><w:ind w:hanging=""400"" w:left=""400"" w:right=""0""/><w:rPr/></w:pPr><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:szCs w:val=""24""/></w:rPr><w:t>"
 Const AbE2$ = "</w:t></w:r><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:szCs w:val=""24""/></w:rPr><w:tab/><w:t>"
 
 Const AbE3$ = "</w:t></w:r><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:szCs w:val=""24""/></w:rPr><w:t xml:space=""preserve""> </w:t><w:tab/><w:t>"
' Const AbE4$ = "•  </w:t></w:r><w:r><w:rPr><w:rFonts w:eastAsia=""Arial"" w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:szCs w:val=""24""/></w:rPr><w:t xml:space=""preserve""> </w:t></w:r><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:szCs w:val=""24""/></w:rPr><w:t>"
 Const AbE4$ = "</w:t><w:t xml:space=""preserve"">&#10151;  </w:t></w:r><w:r><w:rPr><w:rFonts w:eastAsia=""Arial"" w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:szCs w:val=""24""/></w:rPr><w:t xml:space=""preserve""> </w:t></w:r><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:szCs w:val=""24""/></w:rPr><w:t>"
'                </w:t></w:r><w:r><w:rPr><w:rFonts w:eastAsia=""Arial"" w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:szCs w:val=""24""/><w:lang w:val=""en-GB""/></w:rPr><w:t xml:space=""preserve""> </w:t></w:r><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:szCs w:val=""24""/><w:lang w:val=""en-GB""/></w:rPr><w:t>
 Const AbE5$ = "<w:p><w:pPr><w:pStyle w:val=""Normal""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/><w:tab w:val=""left"" w:pos=""3118"" w:leader=""dot""/></w:tabs><w:spacing w:before=""57"" w:after=""0""/><w:ind w:hanging=""400"" w:left=""400"" w:right=""0""/><w:rPr/></w:pPr><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:szCs w:val=""24""/></w:rPr><w:t>"
 Dim ag As New CString
 Dim oSh As New IWshShell_Class
 Call DMPString$(Pat_ID, DT, , True, , False)
' robocopy geht nicht, stellt nicht alle Unterverzeichnisse her, behauptet dann, sie seien schon da
' oSh.rUn "robocopy """ & uVerz & "programmierung\dateilesen\zudocx\\"" """ & pVerz & "zufaxen\" & uvz & "neu"" /s /copy:dat /r:3 /w:2", 0, True
 oSh.rUn "xcopy """ & QuellVerz & "zudocx\*.*"" """ & pVerz & "zufaxen\" & uvz & "neu\"" /s /d /k /y /h /r /c", 0, True
 If FSO.FileExists(dxml) Then Kill dxml
 ag.Append "<w:document xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:wps=""http://schemas.microsoft.com/office/word/2010/wordprocessingShape"" xmlns:wpg=""http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"" xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:wp14=""http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"" xmlns:w15=""http://schemas.microsoft.com/office/word/2012/wordml"" mc:Ignorable=""w14 wp14 w15""><w:body><w:p><w:pPr><w:pStyle w:val=""Heading1""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/>"
 ag.Append "<w:tab w:val=""left"" w:pos=""3118"" w:leader=""dot""/></w:tabs><w:ind w:hanging=""400"" w:left=""400"" w:right=""0""/><w:rPr/></w:pPr><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:sz w:val=""20""/></w:rPr><w:t>Diabetol. Gemeins'praxis Dachau</w:t></w:r><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b w:val=""false""/><w:sz w:val=""20""/></w:rPr><w:t>, G.Schade, Dr.T.Kothny</w:t><w:tab/><w:t>Tel. 08131 616380, Fax 616381"
 ag.Append AbsE
 ag.Append "<w:p><w:pPr><w:pStyle w:val=""Normal""/><w:widowControl w:val=""false""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/><w:tab w:val=""left"" w:pos=""3118"" w:leader=""dot""/></w:tabs><w:autoSpaceDE w:val=""false""/><w:spacing w:lineRule=""exact"" w:line=""100"" w:before=""0"" w:after=""40""/><w:ind w:hanging=""400"" w:left=""400"" w:right=""0""/><w:rPr><w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial"" w:cs=""Arial""/></w:rPr></w:pPr><w:r><w:rPr>"
 ag.AppVar Array("<w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/></w:rPr><w:tab/></w:r></w:p><w:p><w:pPr><w:pStyle w:val=""Normal""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/><w:tab w:val=""left"" w:pos=""3118"" w:leader=""dot""/></w:tabs><w:ind w:hanging=""400"" w:left=""400"" w:right=""0""/><w:rPr><w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial"" w:cs=""Arial""/><w:sz w:val=""18""/><w:szCs w:val=""24""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:sz w:val=""18""/><w:szCs w:val=""24""/></w:rPr><w:t>Sehr geehr. F./H. Kolleg..,</w:t><w:tab/><w:t xml:space=""preserve"">untenst. Pat. ist nach unseren Unterlagen b. Ihnen im DMP Diabetes eingeschrieben. </w:t><w:br/><w:t xml:space=""preserve"">Hier zu Ihrer Unterstützung unsere DMP-Infos. </w:t><w:tab/><w:t>Mit bestem Dank für das Vertrauen und freundlichen Grüßen", AbsE)
 ag.Append "<w:p><w:pPr><w:pStyle w:val=""Normal""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/>"
 ag.AppVar Array("<w:tab w:val=""left"" w:pos=""3118"" w:leader=""dot""/></w:tabs><w:ind w:hanging=""400"" w:left=""400"" w:right=""0""/><w:rPr><w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial"" w:cs=""Arial""/><w:sz w:val=""18""/><w:szCs w:val=""24""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:sz w:val=""18""/><w:szCs w:val=""24""/></w:rPr></w:r></w:p>", AbE1)
' ag.Append replace$(replace$(replace$(DT.x_gesName, "'", "&#8217;"), "ß", "&#223;"), "*", "&#9733;")
 ag.Append sonz(DT.x_GesName)
 ag.Append "</w:t></w:r><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:vanish/><w:szCs w:val=""24""/></w:rPr><w:t xml:space=""preserve""> (Pat'nr "
 ag.Append Pat_ID
 ag.Append ")</w:t></w:r><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:szCs w:val=""24""/></w:rPr><w:t>, DMP-Informatione</w:t></w:r><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:szCs w:val=""24""/></w:rPr><w:t>n für "
 ag.AppVar Array(Format(DT.x_DokuDat, "d.m.yy"), ":", AbsE)
 ag.AppVar Array(AbE1, sonz(DT.x_DmTyp), AbE2, sonz(DT.daseit), AbsE)
 ag.AppVar Array(AbE5, AbE4, "Bek. Begleit-/Folgeerk.:", AbE2, sonz(DT.x_Bbk), AbsE)
 ag.AppVar Array(AbE1, "Raucher:", AbE2, IIf(DT.Tabak, "ja", "nein"), AbsE)
 ag.AppVar Array(AbE1, "Körpergröße:", AbE2, DT.kGR, " cm", AbsE)
 ag.AppVar Array(AbE1, "Körpergewicht:", AbE2, sonz(DT.x_Gewi), AbsE)
 ' <w:p><w:pPr><w:pStyle w:val=""Normal""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/><w:tab w:val=""left"" w:pos=""3118"" w:leader=""dot""/></w:tabs><w:ind w:hanging=""400"" w:left=""400"" w:right=""0""/><w:rPr/></w:pPr><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:szCs w:val=""24""/><w:lang w:val=""en-GB""/></w:rPr><w:t>
 ag.AppVar Array(AbE1, "Blutdruck:", AbE2, sonz(DT.PrRR), AbsE)
 ' <w:p><w:pPr><w:pStyle w:val=""Normal""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/><w:tab w:val=""left"" w:pos=""3118"" w:leader=""dot""/></w:tabs><w:ind w:hanging=""400"" w:left=""400"" w:right=""0""/><w:rPr/></w:pPr><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:szCs w:val=""24""/><w:lang w:val=""en-GB""/></w:rPr><w:t>
 ' </w:t></w:r><w:r><w:rPr><w:rFonts w:eastAsia=""Arial"" w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:szCs w:val=""24""/><w:lang w:val=""en-GB""/></w:rPr><w:t xml:space=""preserve""> </w:t></w:r><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:i/><w:szCs w:val=""24""/><w:lang w:val=""en-GB""/></w:rPr><w:t>
 ' </w:t></w:r><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:szCs w:val=""24""/><w:lang w:val=""en-GB""/></w:rPr><w:tab/><w:t>
 ag.AppVar Array(AbE5, AbE4, "HbA1c:", AbE2, sonz(DT.x_DMPHbA1c), AbsE)
 ag.AppVar Array(AbE1, "eGFR:", AbE2, sonz(DT.x_DMPeGFR), AbsE)
 ag.AppVar Array(AbE1, "Urin-Albumin:", AbE2, sonz(DT.x_DMPUAlb), AbsE)
 ag.AppVar Array(AbE1, "Spätfolgen:", AbE3, sonz(DT.x_Spät), AbsE)
 ag.AppVar Array(AbE1, "Injektionsstellen:", AbE2, sonz(DT.x_Inj), AbsE)
 ag.AppVar Array(AbE5, AbE4, "Pulsstatus:", AbE3, sonz(DT.x_Pulsstatus), AbsE)
 ag.AppVar Array(AbE1, "Sensibilität:", AbE3, sonz(DT.x_SensText), AbsE)
 ag.AppVar Array(AbE1, "Weiteres Risiko für Ulcus:", AbE3, sonz(DT.x_wrstr), AbsE)
 ag.AppVar Array(AbE1, "Ulcus:", AbE3, sonz(DT.x_Ulcus), AbsE)
 ag.AppVar Array(AbE1, "Wundinfektion:", AbE3, sonz(DT.x_Wundinfektion), AbsE)
 ag.AppVar Array(AbE1, "Intervall der künftigen Fußinspektionen:", AbE3, sonz(DT.x_Wundinfektion), AbsE)
 ag.AppVar Array(AbE1, "relev.Ereignisse:", AbE3, sonz(DT.x_Ereig), AbsE)
 ag.AppVar Array(AbE1, "Schw.Hypoglyk./Q.:", AbE3, sonz(DT.x_HypoStr), AbsE)
 If DT.dtyp <> "1" Then
  ag.AppVar Array(AbE5, AbE4, "Insulin oder -analoga:", AbE2, sonz(DT.x_Insulin), AbsE)
  ag.AppVar Array(AbE1, "Metformin:", AbE3, sonz(DT.x_Metformin), AbsE)
  ag.AppVar Array(AbE1, "Sonst.antidiab.Medik.:", AbE3, sonz(DT.x_Sonstige), AbsE)
  ag.AppVar Array(AbE1, "SGLT2-Hemmer:", AbE3, sonz(DT.x_Sglt2), AbsE)
  ag.AppVar Array(AbE1, "GLP1-Analoga:", AbE3, sonz(DT.x_Glp1), AbsE)
  ag.AppVar Array(AbE1, "Thrombozytenhemmer:", AbE3, sonz(DT.x_Thromb), AbsE)
 Else
  ag.AppVar Array(AbE5, AbE4, "Thrombozytenhemmer:", AbE2, sonz(DT.x_Thromb), AbsE)
 End If ' DT.dtyp <> "1" Then
 ag.AppVar Array(AbE1, "Orale Antikoagulation:", AbE3, DT.x_OAK, AbsE)
 ag.AppVar Array(AbE5, AbE4, "Schulungen empfohlen:", AbE2, sonz(DT.x_SchulEmpf), AbsE)
 ag.AppVar Array(AbE1, "Diabetesschulung   wahrg.:", AbE3, sonz(DT.x_SchulWahrg), AbsE)
 ag.AppVar Array(AbE1, "Hypertonieschulung wahrg.:", AbE3, sonz(DT.x_HSchulWahr), AbsE)
 ag.AppVar Array(AbE1, "Bish.Schulg.b.uns (ges.):", AbE3, sonz(DT.x_SchulStr), AbsE)
 ag.AppVar Array(AbE5, AbE4, "Info'angeb.KK gewü.:", AbE2, sonz(DT.x_InfoAng), AbsE)
 ag.AppVar Array(AbE1, "HbA1c-Zielwert:", AbE3, sonz(DT.x_HZiel), AbsE)
 ag.AppVar Array(AbE1, "Bhdlg.Fußeinr.:", AbE3, sonz(DT.x_BhdF), AbsE)
 ag.AppVar Array(AbE1, "Diab'bez.stat.Einweisung:", AbE3, sonz(DT.x_BhdD), AbsE)
 If DT.dtyp = "1" Or Not DT.ab1023 Then ag.AppVar Array(AbE1, "Augenuntersuchung:", AbE3, sonz(DT.x_Aug), AbsE)
 ag.AppVar Array(AbE1, "Empf. Dok'intervall:", AbE3, sonz(DT.x_EmpfItv), AbsE)
' Ende des einspaltigen Teils
 ag.Append "<w:p><w:pPr><w:sectPr><w:type w:val=""nextPage""/><w:pgSz w:w=""11906"" w:h=""16838""/><w:pgMar w:left=""567"" w:right=""567"" w:gutter=""0"" w:header=""0"" w:top=""567"" w:footer=""0"" w:bottom=""567""/><w:pgNumType w:fmt=""decimal""/><w:formProt w:val=""false""/><w:textDirection w:val=""lrTb""/><w:docGrid w:type=""default"" w:linePitch=""360"" w:charSpace=""0""/></w:sectPr></w:pPr></w:p>"

#Else ' wordalt
 Dim dc, VorString$
 Dim mR1, mR2, mR3
' dtbInit
' SET rsAnam = TabÖff("Anamnesebogen", "Pat_id")
' myFrag rsAnam, "SELECT * FROM `anamnesebogen` WHERE pat_id = " & Pat_id ' 22.3.12: nach untens verschoben, da evtl. während des Aufrufs von DMPString Verlorengehen möglich
' rsAnam.Seek "=", Pat_id
 VorString = vNS
 nzw = vbCr
' IF rsAnam!DMPhier <> CDate(0) AND rsAnam!HAimDMP = "Hausarzt im DMP" THEN
'  VorString = "Bei mir fand am " + format$(rsAnam!DMPhier, "dd/mm/yyyy") + " die Einschreibung ins DMP statt. Diese sollte jedoch auf Sie als Hausarzt übertragen werden." + nzw
' END IF
vorgetword:
 GetWord
 With Wapp
  .options.SmartCutPaste = False
  On Error Resume Next
  .options.SmartParaSelection = False
  On Error GoTo fehler
  Call .documents.Add(DMPVorlage)
  Set dc = .activedocument
  Dim VorZahl&, runde%
'  GoTo vorgetword
  FNr = 719
  VorZahl = dc.Range.END
  FNr = 1
 End With ' WApp
 dc.Range.Insertafter VorString & REPLACE$(DMPString$(Pat_ID, DT, , True), vbCrLf & vbCrLf, vbCrLf)
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
   Set mR3 = dc.Range(mR1.Start, mR2.END)
   mR3.Font.Hidden = True
  End If
 End If
        
 Set mR3 = dc.Range(VorZahl, dc.Range.END)
 On Error Resume Next
 Dim Para
 Set Para = mR3.paragraphs.First.Range
 Do While Err.Number = 0
  With dc.Range(Para.Start, Para.Start + InStr(Para.Text, ":")).Font
   .Italic = True
   .bold = True
  End With
  Set Para = Para.paragraphs.First.Next.Range
 Loop
 On Error GoTo fehler
 Dim SectZ%
 SectZ = dc.sections.COUNT
 dc.sections(SectZ).Range.ParagraphFormat.TabStops.ClearAll
 dc.sections(SectZ).Range.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(9#), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderDots
 dc.sections(SectZ).Range.ParagraphFormat.FirstLineIndent = -20
 dc.sections(SectZ).Range.ParagraphFormat.LeftIndent = 20
 dc.Range(dc.Range.END - 1, dc.Range.END - 1).InsertBreak Type:=3 'wdSectionBreakContinuous
 dc.sections.Last.Range.PageSetup.TextColumns.Add Width:=CentimetersToPoints(11.4), Spacing:=CentimetersToPoints(0.6), EvenlySpaced:=False
'  dc.PageSetup.TopMargin = CentimetersToPoints(0.5)
'  dc.PageSetup.BottomMargin = CentimetersToPoints(0.5)

#End If ' not wordalt else
 
 Dim mrs As New ADODB.Recordset
 Set mrs = Nothing
'  sql = "SELECT `medarten`.*, `medplan`.medikament AS mmedikament,bemerkung,mo,mi,nm,ab,zn,--bbed AS j_bbed FROM `medplan` LEFT JOIN `medarten` ON `medplan`.medanfang = `medarten`.medikament WHERE `medplan`.mpnr = (SELECT MAX(mpnr) FROM `medplan` WHERE pat_id = " & Pat_id & " AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM `medplan` WHERE pat_id = " & CStr(Pat_id) & ")) AND NOT ISNULL(`medplan`.medikament) AND `medplan`.medikament <> '' AND `medplan`.pat_id = " & Pat_id
' `medplan`.medikament AS mmedikament,mo,mi,nm,ab,zn
 sql = "SELECT DATE(zeitpunkt) AS zp FROM `medplan` LEFT JOIN `medarten` ON `medplan`.medanfang = `medarten`.medikament WHERE `medplan`.pat_id = " & Pat_ID & " AND `medplan`.mpnr = (SELECT MAX(mpnr) FROM `medplan` WHERE pat_id = " & Pat_ID & " AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM `medplan` WHERE pat_id = " & CStr(Pat_ID) & ")) AND NOT ISNULL(`medplan`.medikament) AND `medplan`.medikament <> '' AND `medplan`.pat_id = " & Pat_ID & " GROUP BY zeitpunkt"
 myFrag mrs, sql
 If Not mrs.BOF Then

#If Not wordalt Then
 ag.Append "<w:p><w:pPr><w:pStyle w:val=""Normal""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/><w:tab w:val=""left"" w:pos=""3118"" w:leader=""dot""/></w:tabs><w:ind w:hanging=""400"" w:left=""400"" w:right=""0""/><w:rPr><w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial"" w:cs=""Arial""/><w:szCs w:val=""24""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:cs=""Arial"" w:ascii=""Arial"" w:hAnsi=""Arial""/><w:szCs w:val=""24""/></w:rPr></w:r></w:p><w:p><w:pPr><w:pStyle w:val=""Normal""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/><w:tab w:val=""left"" w:pos=""3118"" w:leader=""dot""/></w:tabs><w:rPr/></w:pPr><w:r><w:rPr><w:rFonts w:cs=""Courier New"" w:ascii=""Courier New"" w:hAnsi=""Courier New""/><w:b/><w:sz w:val=""16""/><w:szCs w:val=""24""/><w:u w:val=""single""/></w:rPr><w:t>"
 ag.AppVar Array("Med.", Format(mrs!Zp, "d.m.yy"), ":")
 ag.Append "</w:t></w:r>"
 ag.Append "<w:r><w:rPr><w:rFonts w:cs=""Courier New"" w:ascii=""Courier New"" w:hAnsi=""Courier New""/><w:sz w:val=""16""/><w:szCs w:val=""24""/></w:rPr><w:t xml:space=""preserve"">      mo mi nm ab zn </w:t><w:br/>"
 sql = "SELECT CONCAT(LEFT(CONCAT(TRIM(`medplan`.medikament),REPEAT('.',17)),17),LPAD(mo,3,'.'),LPAD(mi,3,'.'),LPAD(nm,3,'.'),LPAD(ab,3,'.'),LPAD(zn,3,'.')) FROM `medplan` LEFT JOIN `medarten` ON `medplan`.medanfang = `medarten`.medikament WHERE `medplan`.Pat_id = " & CStr(Pat_ID) & " AND `medplan`.mpnr = (SELECT MAX(mpnr) FROM `medplan` WHERE pat_id = " & CStr(Pat_ID) & " AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM `medplan` WHERE pat_id = " & CStr(Pat_ID) & ")) AND NOT ISNULL(`medplan`.medikament) AND `medplan`.medikament <> '' AND `medplan`.pat_id = " & CStr(Pat_ID)
 Set mrs = Nothing
 myFrag mrs, sql
 If Not mrs.BOF Then
  Do While Not mrs.EOF
   ag.AppVar Array("<w:t xml:space=""preserve"">", sonz(mrs.Fields(0)), "</w:t><w:br/>")
   mrs.MoveNext
  Loop
 End If ' Not mrs.BOF Then
 ag.Append "</w:r></w:p>"
#Else
  sql = "SELECT LEFT(`medplan`.medikament,16) `Med." & Format(mrs!Zp, "d.m.yy") & ":`,mo,mi,nm,ab,zn FROM `medplan` LEFT JOIN `medarten` ON `medplan`.medanfang = `medarten`.medikament WHERE `medplan`.Pat_id = " & CStr(Pat_ID) & " AND `medplan`.mpnr = (SELECT MAX(mpnr) FROM `medplan` WHERE pat_id = " & CStr(Pat_ID) & " AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM `medplan` WHERE pat_id = " & CStr(Pat_ID) & ")) AND NOT ISNULL(`medplan`.medikament) AND `medplan`.medikament <> '' AND `medplan`.pat_id = " & CStr(Pat_ID)
  Set mrs = Nothing
  myFrag mrs, sql
  dc.Range.Insertafter vbCrLf & TabAusgeb(mrs, Lese, , Chr$(11), , , , , , , True).Value
  Set mR2 = dc.paragraphs.Last.Range
  mR2.Find.Text = ":"
  mR2.Find.Execute
  If mR2.Find.found Then
   dc.Range(dc.paragraphs.Last.Range.Start, mR2.END).bold = True
   dc.Range(dc.paragraphs.Last.Range.Start, mR2.END).Underline = 1
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
  End With ' dc.paragraphs.Last.Range.Font
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
  End With ' dc.paragraphs.Last.Range.ParagraphFormat
#End If ' wordalt
 End If ' Not mrs.BOF Then

 Dim DiagTab() As CString, gesD$
 gesD = DiagString$(CStr(Pat_ID), DiagTab, Now() - 180, True)
#If Not wordalt Then
 ag.Append "<w:p><w:pPr><w:pStyle w:val=""Normal""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/><w:tab w:val=""left"" w:pos=""5102"" w:leader=""dot""/></w:tabs><w:spacing w:before=""170"" w:after=""0""/><w:rPr><w:rFonts w:ascii=""Courier New"" w:hAnsi=""Courier New"" w:cs=""Courier New""/><w:sz w:val=""16""/><w:szCs w:val=""24""/></w:rPr></w:pPr><w:r><w:br w:type=""column""/></w:r><w:r><w:rPr><w:rFonts w:cs=""Courier New"" w:ascii=""Courier New"" w:hAnsi=""Courier New""/><w:b/><w:sz w:val=""16""/><w:szCs w:val=""24""/><w:u w:val=""single""/></w:rPr><w:t>"
 ag.Append "Diagnosen:</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val=""Normal""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/><w:tab w:val=""left"" w:pos=""5102"" w:leader=""dot""/></w:tabs><w:rPr><w:rFonts w:ascii=""Courier New"" w:hAnsi=""Courier New"" w:cs=""Courier New""/><w:sz w:val=""16""/><w:szCs w:val=""24""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:cs=""Courier New"" w:ascii=""Courier New"" w:hAnsi=""Courier New""/><w:sz w:val=""16""/><w:szCs w:val=""24""/></w:rPr>"
 Dim iru&
 For iru = 0 To UBound(DiagTab)
  ag.AppVar Array("<w:t xml:space=""preserve"">", sonz(left$(left$(DiagTab(iru), InStr(DiagTab(iru), Chr$(9)) - 1) & String(45, "."), 45)), "</w:t><w:tab/><w:t>", sonz(Mid$(DiagTab(iru), InStr(DiagTab(iru), Chr(9)) + 1)), "</w:t><w:br/>")
  If iru = 28 Then ' zwei Spaltenumbrüche für die nächste Seite
   ag.Append "</w:r></w:p><w:p><w:pPr><w:pStyle w:val=""Normal""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/><w:tab w:val=""left"" w:pos=""5102"" w:leader=""dot""/></w:tabs><w:rPr><w:rFonts w:ascii=""Courier New"" w:hAnsi=""Courier New"" w:cs=""Courier New""/><w:sz w:val=""16""/><w:szCs w:val=""24""/></w:rPr></w:pPr><w:r><w:br w:type=""column""/></w:r><w:r><w:rPr/></w:r></w:p><w:p><w:pPr><w:pStyle w:val=""Normal""/><w:tabs><w:tab w:val=""clear"" w:pos=""1701""/><w:tab w:val=""left"" w:pos=""5102"" w:leader=""dot""/></w:tabs><w:rPr><w:rFonts w:ascii=""Courier New"" w:hAnsi=""Courier New"" w:cs=""Courier New""/><w:sz w:val=""16""/><w:szCs w:val=""24""/></w:rPr></w:pPr><w:r><w:br w:type=""column""/></w:r><w:r><w:rPr><w:rFonts w:cs=""Courier New"" w:ascii=""Courier New"" w:hAnsi=""Courier New""/><w:sz w:val=""16""/><w:szCs w:val=""24""/></w:rPr>"
  End If
 Next iru
 
 ag.Append "</w:r></w:p>"
 ag.Append "<w:sectPr><w:type w:val=""continuous""/><w:pgSz w:w=""11906"" w:h=""16838""/><w:pgMar w:left=""567"" w:right=""567"" w:gutter=""0"" w:header=""0"" w:top=""567"" w:footer=""0"" w:bottom=""567""/><w:cols w:num=""2"" w:equalWidth=""false"" w:sep=""false""><w:col w:w=""3969"" w:space=""340""/><w:col w:w=""6463""/></w:cols><w:formProt w:val=""false""/><w:textDirection w:val=""lrTb""/><w:docGrid w:type=""default"" w:linePitch=""360"" w:charSpace=""0""/></w:sectPr></w:body></w:document>"
 Open dxml For Output As #197
 Print #197, zsuh(ag.Value)
 Close #197
' docName = pVerz & "zufaxen\" & uvz & DT.x_DoklangName & ".docx"
 docName = docName & "x"
 oSh.rUn "cmd /c """"c:\program files\7-zip\7z"" a -tzip -mm=deflate -mx9 -aoa -r -xr!*.swp """ & docName & """ """ & guv & "*.*""""", 0, True
#Else
 dc.Range(dc.Range.END - 1, dc.Range.END - 1).InsertBreak Type:=8 'wdColumnBreak
 dc.Range(dc.Range.END - 1, dc.Range.END - 1).Insertafter "Diagnosen:" & vbCrLf
 dc.Range(dc.paragraphs.Last.Previous.Range.Start, dc.paragraphs.Last.Previous.Range.END - 1).bold = True
 dc.Range(dc.paragraphs.Last.Previous.Range.Start, dc.paragraphs.Last.Previous.Range.END - 1).Underline = 1 ' wdUnderlineSingle
 dc.Range(dc.Range.END - 1, dc.Range.END - 1).Insertafter gesD
 dc.SaveAs Filename:=REPLACE(REPLACE$(docName, "/", vNS), "\'", "'")
 dc.Close
'  .Visible = True
  
'  .Application.WindowState = wdWindowStateMaximize
'  .Activate
#End If ' wordalt
 Lese.Ausgeb "DMP-Doku erstellt: " & docName, -1
 syscmd 4, "DMP-Doku erstellt: " & docName
 do_DMPAusgebStandAlone = docName
 Exit Function
fehler:
#If wordalt Then
 If FNr = 719 Or FNr = 1 Then
  If runde < 5 Then
   runde = runde + 1
   Resume vorgetword
  End If
 End If ' FNr = 719 Or FNr = 1 Then
#End If
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_DMPAusgebStandAlone/" + AnwPfad)
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
End Function ' ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung

Public Function ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung, 4 x so schnell wie untiges, mit CString hingegen 6 x so langsam wie dieses
Select Case Month(Datum)
 Case 1, 2, 3: ZQuart = "1"
 Case 4, 5, 6: ZQuart = "2"
 Case 7, 8, 9: ZQuart = "3"
 Case Else:    ZQuart = "4"
End Select
ZQuart = ZQuart & Year(Datum)
End Function ' ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung

Public Function ZQKurz$(Datum As Date) ' für PLZ
Select Case Month(Datum)
 Case 1, 2, 3: ZQKurz = "1"
 Case 4, 5, 6: ZQKurz = "2"
 Case 7, 8, 9: ZQKurz = "3"
 Case Else:    ZQKurz = "4"
End Select
ZQKurz = ZQKurz & "/" & Right$(Year(Datum), 2)
End Function ' ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung

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
End Function ' ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung

'Public FUNCTION ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung
'Dim j AS String * 4, q AS String * 1
'On Error GoTo fehler
'j = YEAR(Datum)
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
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZQuart/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung
Public Function zqtest()
 Dim i&, T1#, T2#, var, aktdat As Date
 aktdat = Now()
 T1 = Timer
 For i = 1 To 100000
  var = ZQuart(aktdat)
  aktdat = aktdat - 1
 Next i
 T2 = Timer
' Debug.Print T2 - T1
End Function ' zqtest

' in DMPString
Function GesNamFn$(ByVal rs) ' AS DAO.Recordset) ' s.a. GesName(
 On Error GoTo fehler
   GesNamFn = rs!Nachname & ", " & IIf(IsNull(rs!Titel), vNS, rs!Titel) + IIf(IsNull(rs!Titel) Or LenB(rs!Titel) = 0, vNS, ", ") + rs!Vorname
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GesNam/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GesName

Function dtyp$(Inh$)
Select Case LCase$(Inh)
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
    ElseIf LCase$(rsAnam("Diabetes seit")) = "bu" Then
'      DSeit = format$(rsAnam!Vorgestellt, "mm\/yy")
'      DSeit = format$(Dtb.OpenRecordset("SELECT fanf FROM `faelle` WHERE fid = (SELECT MIN(fid) FROM `faelle` WHERE pat_id = " + CStr(rsAnam!Pat_id) + ")")!Fanf, "mm\/yy")
      myFrag rsDS, "SELECT fanf FROM `faelle` WHERE fid = (SELECT MIN(fid) FROM `faelle` WHERE pat_id = " + CStr(rsAnam!Pat_ID) + ")"
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DSeit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DSeit(rsAnam)

Function TabPr(s1, s2)
'Dim nzl$
'nzl = vbcr + vblf ' vbcrlf kannte ich da noch nicht
#If False Then
 DmPStrS.AppVar Array(vbCrLf, s1)
 If Len(s1) <= BreiteSp1 Then
  DmPStrS.AppVar Array(left$(" ........................................", BreiteSp1 - Len(s1) - 1), " ")
 End If
 DmPStrS.Append IIf(IsNull(s2), vNS, s2) ' "........................"
#Else
 DmPStrS.AppVar Array(vbCrLf, s1, Chr$(9), IIf(IsNull(s2), vNS, s2))
#End If
End Function ' TabPr(S1, S2)

'obpositiv
' in DMPString
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
 End If ' isnull(s) else else
End Function ' obPosi(S) AS Boolean

Public Function neuTher()
 Dim rNa As New ADODB.Recordset
 Call Lese.ProgStart
 myFrag rNa, "SELECT pat_id,gesname(pat_id) name FROM namen"
 Do While Not rNa.EOF
'  Debug.Print rNa!Pat_id, rNa!name
  myEFrag "SELECT therartn(" & rNa!Pat_ID & ")"
  rNa.MoveNext
  DoEvents
 Loop
End Function ' neuTher()

#If True Then
Function TherAuskunft(ByVal Pat_ID$, ByVal obanf%, Optional ByRef insz%, Optional ByVal VorDat As Date, _
 Optional ByRef obIns As Boolean, Optional ByRef obAnal As Boolean, Optional ByRef obGlib As AntidiabMedType, _
 Optional ByRef obmetf As AntidiabMedType, Optional ByRef obGlucI As AntidiabMedType, _
 Optional ByRef obSHGlin As AntidiabMedType, Optional ByRef obGlit As AntidiabMedType, _
 Optional ByRef obDpp4 As AntidiabMedType, Optional ByRef obGlp1 As AntidiabMedType, _
 Optional ByRef obSglt2 As AntidiabMedType, Optional ByRef obSonstAD As AntidiabMedType, _
 Optional ByRef obHMG As Boolean, Optional ByRef obAntihyp As Boolean, Optional ByRef obACEH As Boolean, _
 Optional ByRef obBetabl As Boolean, Optional ByRef obThro As Boolean, Optional ByRef obAntikoag%, _
 Optional ByVal Qmax$, Optional ByVal obRezIns As Boolean, Optional ByVal DT$, Optional ByRef obDiur%, _
 Optional ByRef obAT1%) As TherapieArt

 Dim raMa As New ADODB.Recordset
 Dim rAf&
 Dim MPNr&(), MP0&, MPe&, mpz&, i&
 Dim sqlIns$, sqlMain$, sqlDosH$, mpList$
 On Error GoTo fehler
 Call Lese.ProgStart

 mpz = 0
 ReDim MPNr(mpz)

 If Qmax <> vNS Then
  MP0 = MedPlanNr(Pat_ID, obAkt:=True, VorDat:=fctQAnf(Qmax), NurNr:=True)
  MPe = MedPlanNr(Pat_ID, obAkt:=True, VorDat:=fctQEnd(Qmax) + 1, NurNr:=True)
  Set raMa = Nothing
  myFrag raMa, "SELECT DISTINCT mpnr FROM `medplan` WHERE pat_id = " & Pat_ID & " AND mpnr >= " & MP0 & " AND mpnr <= " & MPe & " ORDER BY mpnr"
  Do While Not raMa.EOF
   mpz = mpz + 1
   ReDim Preserve MPNr(mpz)
   MPNr(mpz) = raMa!MPNr
   raMa.Move 1
  Loop
 Else
  mpz = 1
  ReDim MPNr(mpz)          ' Index 1 statt 0
  If obanf Then
   Call MedPlanAusAna(Pat_ID)
   MPNr(1) = 0             ' MedPlanAusAna befüllt Med/Dos direkt, MPNr irrelevant
  ElseIf obanf = 1 Then
   MPNr(1) = MedPlanNr(Pat_ID, -1)
  Else
   MPNr(1) = MedPlanNr(Pat_ID, Not obanf, VorDat)
  End If
 End If

 ' DosH-Ausdruck: mind. ein Dosierungsfeld non-empty und nicht "0"
 ' Entspricht DosHäuf > 0
 sqlDosH = "(CASE WHEN " & _
  "COALESCE(NULLIF(NULLIF(TRIM(mp.mo),''),'0'), " & _
  "NULLIF(NULLIF(TRIM(mp.mi),''),'0'), " & _
  "NULLIF(NULLIF(TRIM(mp.nm),''),'0'), " & _
  "NULLIF(NULLIF(TRIM(mp.ab),''),'0'), " & _
  "NULLIF(NULLIF(TRIM(mp.zn),''),'0')) IS NOT NULL THEN 1 ELSE 0 END)"

 For i = 1 To mpz
  If Qmax <> "" Then
   Call MedPlanNr(Pat_ID, True, , , nr:=MPNr(i))
  End If

  obGlib = adnein: obmetf = adnein: obGlucI = adnein: obSHGlin = adnein: obGlit = adnein
  obDpp4 = adnein: obGlp1 = adnein: obSglt2 = adnein: obSonstAD = adnein
  obIns = 0: obAnal = 0: obHMG = 0: obAntihyp = 0: obACEH = 0: obBetabl = 0: obThro = 0
  insz = 0

  ' --- 1. INSERT für unbekannte Medikamente mit DosH > 0 ---
  ' Alle MedAnfang-Werte aus medplan für diese MPNr, die noch nicht in medarten stehen
  sqlIns = "INSERT IGNORE INTO `medarten` (langname, medikament, hinzugefügt, pat_id) " & _
   "SELECT DISTINCT mp.medikament, mp.medanfang, NOW(), mp.pat_id " & _
   "FROM `medplan` mp " & _
   "WHERE mp.pat_id = " & Pat_ID & " AND mp.mpnr = " & MPNr(i) & _
   " AND " & sqlDosH & " > 0 " & _
   " AND mp.medanfang IS NOT NULL AND TRIM(mp.medanfang) <> '' AND TRIM(mp.medanfang) <> '-'" & _
   " AND NOT EXISTS (SELECT 1 FROM `medarten` ma WHERE ma.medikament = mp.medanfang)"
  InsKorr DBCn, sqlIns, rAf
  ' (Debug-Ausgabe wie bisher entfällt, da INSERT IGNORE das still erledigt;
  '  wer das logging braucht, kann vorher per SELECT prüfen)

  ' --- 2. Hauptabfrage: Aggregation über alle Medis des Plans ---
    ' kein DosH-Filter!
' Vorab als Konstante definieren (oder in der SQL inline lassen):
 Dim sqlGlp1Bem$
 sqlGlp1Bem = _
  "((TRIM(COALESCE(mp.bemerkung,'')) = '' AND mp.grund RLIKE '[0-9 ]*[ck]licks|jeden|einmal|Diabetes|zucker|gewicht')" & _
  " OR mp.bemerkung RLIKE 'nachmittags|gesteigert|every|jd\\.|jed|Woche|wöch|[0-9 ]*[ck]licks|^[-=]+>|alle| Dj$|^([0-9]|Dg\\.|ab |am |Beginn|Dj |Mo|Di|Mi|Do|Fr|Sa|So|statt|wenn|siehe|seit|dann|danach|D:|ca\\.|bitte gleich|immer|nur|once|neu|n[äa]|erhöht|reduziert|ersatzweise|nach |Wiederbeginn|Privat|Umstellung)|^einmal|^D: |^aufbrauch|^als n[äa]|1x|Lieferengpass|vorübergehend|erhöht')"
  sqlMain = _
   "SELECT " & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.glib,0) <> 0 THEN 1 ELSE 0 END)  AS ag_glib," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.metf,0) <> 0 THEN 1 ELSE 0 END)  AS ag_metf," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.gluci,0) <> 0 THEN 1 ELSE 0 END) AS ag_gluci," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.shglin,0) <> 0 THEN 1 ELSE 0 END) AS ag_shglin," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.glit,0) <> 0 THEN 1 ELSE 0 END)  AS ag_glit," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.dpp4,0) <> 0 THEN 1 ELSE 0 END)  AS ag_dpp4," & _
   " MAX(CASE WHEN COALESCE(ma.glp1,0) <> 0 AND (" & sqlDosH & " > 0 OR " & sqlGlp1Bem & ") THEN 1 ELSE 0 END) AS ag_glp1," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.sglt2,0) <> 0 THEN 1 ELSE 0 END) AS ag_sglt2," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.sonstad,0) <> 0 THEN 1 ELSE 0 END) AS ag_sonstad," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND (COALESCE(ma.ins,0) <> 0 OR (COALESCE(ma.anal,0) = 0 AND mp.medanfang LIKE '%INSULIN%')) THEN 1 ELSE 0 END) AS ag_ins," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.anal,0) <> 0 THEN 1 ELSE 0 END)  AS ag_anal," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.insart,0) = 1 THEN 1 ELSE 0 END)  AS ag_nins," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.hmg,0) <> 0 THEN 1 ELSE 0 END)   AS ag_hmg," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.hypt,0) <> 0 THEN 1 ELSE 0 END)  AS ag_hypt," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.diur,0) <> 0 THEN 1 ELSE 0 END)  AS ag_diur," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.betabl,0) <> 0 THEN 1 ELSE 0 END) AS ag_betabl," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.aceh,0) <> 0 THEN 1 ELSE 0 END)  AS ag_aceh," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.at1,0) <> 0 THEN 1 ELSE 0 END)   AS ag_at1," & _
   " MAX(CASE WHEN " & sqlDosH & " > 0 AND COALESCE(ma.thro,0) <> 0 THEN 1 ELSE 0 END)  AS ag_thro,"
  sqlMain = sqlMain & _
   " MAX(CASE WHEN COALESCE(ma.antikoag,0) <> 0 THEN 1 ELSE 0 END) AS ag_antikoag," & _
   " SUM(CASE WHEN " & sqlDosH & " > 0 AND (COALESCE(ma.ins,0) <> 0 OR COALESCE(ma.anal,0) <> 0 OR (COALESCE(ma.ins,0) = 0 AND COALESCE(ma.anal,0) = 0 AND mp.medanfang LIKE '%INSULIN%')) " & _
   "     THEN " & sqlDosH & " ELSE 0 END) AS ag_insz" & _
   " FROM `medplan` mp" & _
   " LEFT JOIN `medarten` ma ON ma.medikament = mp.medanfang" & _
   " WHERE mp.pat_id = " & Pat_ID & " AND mp.mpnr = " & MPNr(i) & _
   " AND mp.medanfang IS NOT NULL AND TRIM(mp.medanfang) <> '' AND TRIM(mp.medanfang) <> '-'"

  Set raMa = Nothing
  myFrag raMa, sqlMain

  Dim obNIns As Boolean
  If Not raMa.BOF And Not raMa.EOF Then
   If nz(raMa!ag_antikoag, 0) <> 0 Then obAntikoag = True
   If nz(raMa!ag_glib, 0) <> 0 Then obGlib = adja
   If nz(raMa!ag_metf, 0) <> 0 Then obmetf = adja
   If nz(raMa!ag_gluci, 0) <> 0 Then obGlucI = adja
   If nz(raMa!ag_shglin, 0) <> 0 Then obSHGlin = adja
   If nz(raMa!ag_glit, 0) <> 0 Then obGlit = adja
   If nz(raMa!ag_dpp4, 0) <> 0 Then obDpp4 = adja
   If nz(raMa!ag_glp1, 0) <> 0 Then obGlp1 = adja
   If nz(raMa!ag_sglt2, 0) <> 0 Then obSglt2 = adja
   If nz(raMa!ag_sonstad, 0) <> 0 Then obSonstAD = adja
   If nz(raMa!ag_ins, 0) <> 0 Then obIns = True
   If nz(raMa!ag_anal, 0) <> 0 Then obAnal = True
   obNIns = (nz(raMa!ag_nins, 0) <> 0)
   If nz(raMa!ag_hmg, 0) <> 0 Then obHMG = True
   If nz(raMa!ag_hypt, 0) <> 0 Then obAntihyp = True
   If nz(raMa!ag_diur, 0) <> 0 Then obDiur = True
   If nz(raMa!ag_betabl, 0) <> 0 Then obBetabl = True
   If nz(raMa!ag_aceh, 0) <> 0 Then obACEH = True
   If nz(raMa!ag_at1, 0) <> 0 Then obAT1 = True
   If nz(raMa!ag_thro, 0) <> 0 Then obThro = True
   insz = nz(raMa!ag_insz, 0)
  End If
  
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
       If obGlp1 = adja Then
        Thfakt = glp1
       ElseIf obGlib = adja Or obmetf = adja Or obGlucI = adja Or obSHGlin = adja Or obGlit = adja Or obDpp4 = adja Or obSglt2 = adja Or obSonstAD = adja Then
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
      If obGlp1 = adja Then
       Thfakt = glp1ins
      ElseIf obGlib = adja Or obmetf = adja Or obGlucI = adja Or obSHGlin = adja Or obGlit = adja Or obDpp4 = adja Or obSglt2 = adja Or obSonstAD = adja Then
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

Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in TherAuskunft/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function
#Else

' wird in TherArt() aufgerufen, wo anhand des Diabetestyps und von Formularen Pumpentherapien ermittelt werden, sodann hiermit die übrigen Therapieformen
' parallel zu ThaFestleg
' aufgerufen in DMPString, Epikrise (2x) und Labordateianzeig
Function TherAuskunft(ByVal Pat_ID$, ByVal obanf%, Optional ByRef insz%, Optional ByVal VorDat As Date, _
 Optional ByRef obIns As Boolean, Optional ByRef obAnal As Boolean, Optional ByRef obGlib As AntidiabMedType, _
 Optional ByRef obmetf As AntidiabMedType, Optional ByRef obGlucI As AntidiabMedType, _
 Optional ByRef obSHGlin As AntidiabMedType, Optional ByRef obGlit As AntidiabMedType, _
 Optional ByRef obDpp4 As AntidiabMedType, Optional ByRef obGlp1 As AntidiabMedType, _
 Optional ByRef obSglt2 As AntidiabMedType, Optional ByRef obSonstAD As AntidiabMedType, _
 Optional ByRef obHMG As Boolean, Optional ByRef obAntihyp As Boolean, Optional ByRef obACEH As Boolean, _
 Optional ByRef obBetabl As Boolean, Optional ByRef obThro As Boolean, Optional ByRef obAntikoag%, _
 Optional ByVal Qmax$, Optional ByVal obRezIns As Boolean, Optional ByVal DT$, Optional ByRef obDiur%, _
 Optional ByRef obAT1%) As TherapieArt
' Dim rMA AS DAO.Recordset
 Dim obNIns% 'ob Normalinsulin
 Dim raMa As New ADODB.Recordset
 Dim AktMz%, medi$, rAf&
 Dim MPNr&(), MP0&, MPe&, mpz&, i&, MPNrl& ' erste und letzte, laufende MedPlan-Nummer des Quartals
 On Error GoTo fehler
 Call Lese.ProgStart
' QuartalMax: Falls befüllt, so wird die höchstwertige Therapie des in der Variablen angegebenen Quartals ermittelt
' SET rMa = TabÖff("MedArten", "Medikament")
'Medikamente:
 mpz = 0
 ReDim MPNr(mpz)
 If Qmax <> vNS Then
  MP0 = MedPlanNr(Pat_ID, obAkt:=True, VorDat:=fctQAnf(Qmax), NurNr:=True)
  MPe = MedPlanNr(Pat_ID, obAkt:=True, VorDat:=fctQEnd(Qmax) + 1, NurNr:=True)
  Set raMa = Nothing
  myFrag raMa, "SELECT DISTINCT mpnr FROM `medplan` WHERE pat_id = " & Pat_ID & " AND mpnr >= " & MP0 & " AND mpnr <= " & MPe & " ORDER BY mpnr"
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
   Call MedPlanAusAna(Pat_ID)
  ElseIf obanf = 1 Then
   MPNr(0) = MedPlanNr(Pat_ID, -1)
  Else
   MPNr(0) = MedPlanNr(Pat_ID, Not obanf, VorDat)
  End If
 End If
 
 For i = 1 To mpz 'MPNrl = MPNrA To MPNrE
  If Qmax <> "" Then
   Call MedPlanNr(Pat_ID, True, , , nr:=MPNr(i)) ' dann die übrigen Variablen hier jeweils neu befüllen
  End If
  obGlib = adnein: obmetf = adnein: obGlucI = adnein: obSHGlin = adnein: obGlit = adnein: obDpp4 = adnein: obGlp1 = adnein: obSglt2 = adnein: obSonstAD = adnein
  obIns = 0: obAnal = 0: obHMG = 0: obAntihyp = 0: obACEH = 0: obBetabl = 0: obThro = 0
  insz = 0
  obNIns = 0
  For AktMz = 0 To MedZahl
   Dim DosH%
   DosH = DosHäuf(AktMz)
'  IF DosH > 0 THEN ' wenn auch bei Dosierung was drinsteht
   medi = Trim$(LCase$(Med(AktMz)))
   If LenB(medi) <> 0 And medi <> "-" Then
'    dim lzpos%
'    lzpos = InStr(medi, " ")
'    IF lzpos > 0 THEN medi = LEFT(medi, lzpos - 1)
    medi = GetMed(medi, 0)
'    rMA.Seek "=", medi
    Set raMa = Nothing
'    raMa.Open "SELECT -hmg j_hmg, -hypt j_hypt, -aceh j_aceh, -betabl j_betabl, -thro j_thro, -antikoag j_antikoag, -glib j_glib, -metf j_metf, -gluci j_gluci, -shglin j_shglin, -glit j_glit, -dpp4 j_dpp4, -glp1 j_glp1, -sglt2 j_sglt2, -sonstad j_sonstad, -ins j_ins, -AnAl j_anal, -Diur j_diur, -AT1 j_at1, m.* FROM `medarten` m WHERE medikament = '" & replace$(medi, "'", "''") & "'", DBCn, adOpenDynamic, adLockReadOnly
    myFrag raMa, "SELECT -hmg j_hmg, -hypt j_hypt, -aceh j_aceh, -betabl j_betabl, -thro j_thro, -antikoag j_antikoag, -glib j_glib, -metf j_metf, -gluci j_gluci, -shglin j_shglin, -glit j_glit, -dpp4 j_dpp4, -glp1 j_glp1, -sglt2 j_sglt2, -sonstad j_sonstad, -ins j_ins, -AnAl j_anal, -Diur j_diur, -AT1 j_at1, m.* FROM `medarten` m WHERE medikament = '" & REPLACE$(medi, "'", "''") & "'", adOpenDynamic
    If raMa.BOF Then
     If DosH > 0 Then
'      MsgBox "Medikament: " + Medi + " noch nicht in Tabelle MedArten erfaßt"
      Debug.Print "Medikament: " + medi + " noch nicht in Tabelle `medarten` erfaßt"
      InsKorr DBCn, "INSERT INTO `medarten`(langname,medikament,hinzugefügt,pat_id) VALUES('" & Med(AktMz) & "','" & UCase$(medi) & "'," & DatFor_k(Now) & "," & Pat_ID & ")", rAf
     End If
    Else
     If Not IsNull(raMa!j_antikoag) Then If raMa!j_antikoag <> 0 Then obAntikoag = True ' Bei Marcumar steht meist keine Dosierung drin
     If DosH > 0 Then
      If Not IsNull(raMa!j_glib) Then If raMa!j_glib <> 0 Then obGlib = adja
      If Not IsNull(raMa!j_Metf) Then If raMa!j_Metf <> 0 Then obmetf = adja
      If Not IsNull(raMa!j_GlucI) Then If raMa!j_GlucI <> 0 Then obGlucI = adja
      If Not IsNull(raMa!j_SHGlin) Then If raMa!j_SHGlin <> 0 Then obSHGlin = adja
      If Not IsNull(raMa!j_Glit) Then If raMa!j_Glit <> 0 Then obGlit = adja
      If Not IsNull(raMa!j_dpp4) Then If raMa!j_dpp4 <> 0 Then obDpp4 = adja
      If Not IsNull(raMa!j_glp1) Then If raMa!j_glp1 <> 0 Then obGlp1 = adja
      If Not IsNull(raMa!j_sglt2) Then If raMa!j_sglt2 <> 0 Then obSglt2 = adja
      If Not IsNull(raMa!j_SonstAD) Then If raMa!j_SonstAD <> 0 Then obSonstAD = adja
      If Not IsNull(raMa!j_InS) And (raMa!j_InS <> 0 Or (Not raMa!j_AnAl <> 0 And InStrB(LCase$(medi), "insulin") > 0)) Then
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
'  END IF 'doshäuf(aktmz)
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
       If obGlp1 = adja Then
        Thfakt = glp1
       ElseIf obGlib = adja Or obmetf = adja Or obGlucI = adja Or obSHGlin = adja Or obGlit = adja Or obDpp4 = adja Or obSglt2 = adja Or obSonstAD = adja Then
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
      If obGlp1 = adja Then
       Thfakt = glp1ins
      ElseIf obGlib = adja Or obmetf = adja Or obGlucI = adja Or obSHGlin = adja Or obGlit = adja Or obDpp4 = adja Or obSglt2 = adja Or obSonstAD = adja Then
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
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in TherAuskunft/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' TherAuskunft(pat_id, obanf%, insz, Optional obIns, Optional obAnal, Optional obGlib, Optional obMetf, Optional obGlucI, Optional obSHGlin, Optional obGlit, Optional obSonstAD, Optional obHMG, Optional obAntihyp, Optional obThro)
#End If

Function diI%(icd_str$, Optional pid, Optional abDat As Date, Optional auchZ As Boolean, Optional Weiteres$, Optional ByRef ICD$, Optional ByRef diagdat As Date)
 Dim raDT As New ADODB.Recordset, raFa As New ADODB.Recordset
 Dim diIsql$
 On Error GoTo fehler
 If Not IsMissing(pid) Then Pat_ID = CLng(pid)
 If Pat_ID = 0 Then
  MsgBox "Stop in diI durch Pat_id = 0"
  Stop
  Exit Function
 End If
 Dim icdsql$
 If InStrB(icd_str, " ") = 0 Then
  icdsql = "icd LIKE '" & icd_str & "%'"
 Else
 ' ergänzt 13.8.15
  Dim SpliZ&, i&, splerg$()
  SpliZ = SplitNeu(icd_str, " ", splerg)
  icdsql = "("
  For i = 0 To SpliZ - 1
   icdsql = icdsql & "icd LIKE '" & splerg(i) & "%'"
   If i = SpliZ - 1 Then icdsql = icdsql & ")" Else icdsql = icdsql & " OR "
  Next i
 End If
 diIsql = "SELECT " & IIf(lies.obMySQL, vNS, "top 1 ") & "* FROM `diagnosen` WHERE pat_id = " & Pat_ID & " AND " & icdsql & " AND diagsicherheit IN ('G','V',''" & IIf(auchZ, ",'Z'", "") & ") " & IIf(abDat <> 0, " AND diagdatum >= " & DatFor_k(abDat), "") & IIf(LenB(Weiteres) <> 0, " AND " & Weiteres, "") '  AND COALESCE(Dggel,0)=0
 Dim lddat As Date
 myFrag raFa, "SELECT MAX(bhfb) AS lddat FROM `faelle` WHERE pat_id = " & Pat_ID
 If Not raFa.BOF Then
  lddat = IIf(IsNull(raFa!lddat), 0, raFa!lddat)
  diIsql = diIsql & " AND (obdauer <> 0 OR diagdatum >= " & DatFor_k(lddat) & ")"
 End If
 If lies.obMySQL Then diIsql = diIsql & " LIMIT 1"
 myFrag raDT, diIsql
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
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diI/" + AnwPfad)
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
 If Not IsMissing(pid) Then Pat_ID = pid
 If Pat_ID = 0 Then
  MsgBox "Stop in diT durch Pat_id = 0"
  Stop
  Exit Function
 End If
 Dim dtxtsql$
 If InStrB(diagtxt, " ") = 0 Then
  dtxtsql = "diagtext LIKE '%" & diagtxt & "%'"
 Else
 ' ergänzt 13.8.15
  Dim SpliZ&, i&, splerg$()
  SpliZ = SplitNeu(diagtxt, " ", splerg)
  dtxtsql = "("
  For i = 0 To SpliZ - 1
   dtxtsql = dtxtsql & "diagtext LIKE '%" & splerg(i) & "%'"
   If i = SpliZ - 1 Then dtxtsql = dtxtsql & ")" Else dtxtsql = dtxtsql & " OR "
  Next i
 End If
 diTsql = "SELECT " & IIf(lies.obMySQL, vNS, "top 1 ") & "* FROM `diagnosen` WHERE pat_id = " & Pat_ID & " AND " & dtxtsql & " AND diagsicherheit IN ('G','V',''" & IIf(auchZ, ",'Z'", "") & ") " & IIf(abDat <> 0, " AND diagdatum >= " & DatFor_k(abDat), "") & IIf(LenB(Weiteres) <> 0, " AND " & Weiteres, "") ' AND COALESCE(Dggel,0)=0
 Dim lddat As Date
 myFrag raFa, "SELECT MAX(bhfb) lddat FROM `faelle` WHERE pat_id = " & Pat_ID
 If Not raFa.BOF Then
  If Not IsNull(raFa!lddat) Then
   lddat = raFa!lddat
   diTsql = diTsql & " AND (obdauer <> 0 OR diagdatum >= " & DatFor_k(lddat) & ")"
  End If
 End If
 If lies.obMySQL Then diTsql = diTsql & " LIMIT 1"
 myFrag raDT, diTsql
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
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diT/" + AnwPfad)
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
' END SELECT
'End FUNCTION ' seite$()

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
     If InStrB(Add, "/") = 0 Then Add = REPLACE$(Add, ",", "/")
     Add = REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(REPLACE$(Add, "biph.", "biphasisch "), "biphas.", "biphasisch "), "hyperä.", "hyperämisch "), "hyperä", "hyperämisch "), "mono", "monophasisch"), "pst", "poststenotisch ")
     If Right$(Add, 2) = "bi" Then Add = Add + "phasisch"
     If InStrB(Add, ",") = 0 And InStrB(Add, "/") = 0 And InStrB(Add, "bds") = 0 And InStrB(Add, "beid") = 0 Then
      Add = "bds. " + Add
     End If
     PulsParse = Add
     rE = Add
     li = Add
     If InStrB(Add, "/") <> 0 Then
      rE = left$(Add, InStr(Add, "/") - 1)
      li = Mid$(Add, InStr(Add, "/") + 1)
     End If
     If InStrB(Add, "|") <> 0 Then
      rE = left$(Add, InStr(Add, "|") - 1)
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
 End If ' isnull(s) else
End Function ' TFeld$(S)

Function obMonPath(Feld, rep%, lip%)
Dim TestFeld$, Tre$, Tli$
rep = 1
lip = 1
SensSplit Feld, Tre, Tli
If InStrB(Tre, "/5") <> 0 Then
 Tre = REPLACE$(Tre, "/5", "")
 Tli = REPLACE$(Tli, "/5", "")
 If Tre = "" Or Tre = "~" Or InStrB(Tre, "5") <> 0 Then rep = 0
 If Tli = "" Or Tli = "~" Or InStrB(Tli, "5") <> 0 Then lip = 0
ElseIf InStrB(Tre, "/3") <> 0 Then
 Tre = REPLACE$(Tre, "/3", "")
 Tli = REPLACE$(Tli, "/3", "")
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
Tre = REPLACE$(Tre, "/5", "")
Tli = REPLACE$(Tli, "/5", "")
If Tre = "" Or Tre = "~" Or InStrB(Tre, "4") <> 0 Or InStrB(Tre, "5") <> 0 Then rep = 0
If Tli = "" Or Tli = "~" Or InStrB(Tli, "4") <> 0 Or InStrB(Tli, "5") <> 0 Then lip = 0
End Function ' obKWPath(Feld, rep%, lip%)
Function obVibPath(Feld, rep%, lip%, Optional obStreng%)
Dim TestFeld$, Tre$, Tli$
rep = 1
lip = 1
SensSplit Feld, Tre, Tli
Tre = REPLACE$(Tre, "/8", "")
Tli = REPLACE$(Tli, "/8", "")
If (Tre = "" Or Tre = "~" Or (obStreng = 0 And InStrB(Tre, "4") <> 0) Or InStrB(Tre, "5") <> 0 Or InStrB(Tre, "6") <> 0 Or InStrB(Tre, "7") <> 0 Or InStrB(Tre, "8") <> 0) And InStrB(Tre, ">8") = 0 Then rep = 0
If (Tli = "" Or Tli = "~" Or (obStreng = 0 And InStrB(Tli, "4") <> 0) Or InStrB(Tli, "5") <> 0 Or InStrB(Tli, "6") <> 0 Or InStrB(Tli, "7") <> 0 Or InStrB(Tli, "8") <> 0) And InStrB(Tli, ">8") = 0 Then lip = 0
End Function ' obVibPath(Feld, rep%, lip%)

Function SensSplit(SensStrg, rE$, li$)
If IsNull(SensStrg) Then SensStrg = ""
If InStrB(SensStrg, "|") <> 0 Then
      rE = LTrim$(left$(CStr(SensStrg), InStr(CStr(SensStrg), "|") - 1))
      li = LTrim$(Mid$(CStr(SensStrg), InStr(CStr(SensStrg), "|") + 1))
ElseIf InStrB(SensStrg, ",") <> 0 Then
      rE = left$(CStr(SensStrg), InStr(CStr(SensStrg), ",") - 1)
      li = Mid$(CStr(SensStrg), InStr(CStr(SensStrg), ",") + 1)
Else
      rE = SensStrg
      li = SensStrg
End If
rE = Trim$(rE)
li = Trim$(li)
End Function ' SensSplit(SensStrg, re$, li$)

' in Epikrise, rrpruef und DMPString
 Public Function GetPrRR$(ByVal pid$, rsAnahier As ADODB.Recordset, RRsyst%, RRdiast%, Optional obdiastkorr% = 0)
 '  Dim rrr AS DAO.Recordset
'  SET rrr = TabÖff("RR", "Auswahl")
'  rrr.Seek "=", CStr(rs!Pat_id)
  On Error GoTo fehler
  GetPrRR = vNS
'  rarr.Open "SELECT * FROM rr WHERE pat_id = " & rsAnahier!Pat_id & " ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
  Dim rarr As New ADODB.Recordset
'  rarr.Open "SELECT rs,rd FROM (SELECT rrsyst(rr) rs, rrdiast(rr) rd, zeitpunkt FROM rr WHERE Pat_id = " & rsAnahier!Pat_id & ") i WHERE rs<>0 AND rd<>0 ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
  myFrag rarr, "SELECT rs,rd FROM (SELECT rrsyst(rr) rs, rrdiast(rr) rd, zeitpunkt FROM rr WHERE Pat_id = " & pid & ") i WHERE rs<>0 AND rd<>0 ORDER BY zeitpunkt DESC"
  If rarr.EOF Then
   Set rarr = Nothing
'   rarr.Open "SELECT rs,rd FROM (SELECT rrsyst(rr) rs, rrdiast(rr) rd, zeitpunkt FROM rr WHERE Pat_id = " & rsAnahier!Pat_id & ") i WHERE rs<>0 ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
   myFrag rarr, "SELECT rs,rd FROM (SELECT rrsyst(rr) rs, rrdiast(rr) rd, zeitpunkt FROM rr WHERE Pat_id = " & pid & ") i WHERE rs<>0 ORDER BY zeitpunkt DESC"
  End If
  If Not rarr.BOF Then
   RRsyst = rarr!rs
   RRdiast = rarr!rD
   GetPrRR = RRsyst & "/" & RRdiast & " mm Hg"
  Else
   If rsAnahier.State = 0 Then
    Set rsAnahier = Nothing
    myFrag rsAnahier, "SELECT * FROM anamnesebogen WHERE pat_id = " & pid, adOpenStatic
   End If
   With rsAnahier
'    IF Not rarr.EOF THEN
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
  If obdiastkorr Then If RRdiast < 30 Then RRdiast = MINvb(70, RRsyst) ' 31.10.20
  Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetPrRR/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GetPrRR

#If False Then
Function SchulzBest%(Pat_ID&, zp1$, zpl$, Optional abDat)
  Dim sqls$, lapp As New ADODB.Recordset
  On Error GoTo fehler
  sqls$ = "SELECT * FROM `eintraege` WHERE pat_id = " & CStr(Pat_ID) + " AND art = ""schul"" "
  If Not IsMissing(abDat) Then If Now > CDate("15.10.05") Then sqls = sqls + "and Zeitpunkt >= " & DatFor_k(abDat) & " "
  sqls = sqls + "ORDER BY zeitpunkt"
'  lapp.Open sqls, DBCn, adOpenDynamic, adLockReadOnly
  myFrag lapp, sqls
  If Not lapp.BOF Then
   lapp.MoveLast
   zpl = Format$(lapp!Zeitpunkt, "dd/mm/yy")
   lapp.MoveFirst
   zp1 = Format$(lapp!Zeitpunkt, "dd/mm/yy")
  End If
  lapp.Close
  myFrag lapp, "SELECT COUNT(0) AS zl FROM (" & sqls & ")" & IIf(lies.obMySQL, " AS innen", "")
  SchulzBest = lapp!zl
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SchutzBest/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'SchulzBest
#End If

Function rrEmpf$(RRsyst%, RRdiast%, Pat_ID$, Optional freq%)
  Dim gwSyst%, gwDiast%, Verzei$, Datei$, taskid&
  Dim erg$
  Call Lese.ProgStart
  If Not diI("I10", Pat_ID) Then
   If RRsyst > 149 Or RRdiast > 100 Then
'    IF MsgBox("Achtung: Fehlt Diagnose Hypertonie bei Pat_id = " & CStr(Pat_id) & "?", vbYesNo) = vbYes THEN
'     MsgBox "Bitte in Turbomed eintragen"
'    END IF
    Verzei = pVerz & "plz\tmp\"
    Datei = Verzei & "Hypertoniehinweis Pat." & Pat_ID & ".txt"
    VerzPrüf Verzei
    erg = Dir(Datei)
    If LenB(erg) = 0 Then
     Open Datei For Output As #371
     Print #371, "Pat." & Pat_ID & " hat RR eingetragen von: " & RRsyst & "/" & RRdiast & "."
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
    If diI("N08.3", Pat_ID) And Not obkNeph Then
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
 Dim buch As String * 1, SL%, aktb%, p1%, pl%, puZ%, ZwiStr$
 SL = Len(teststr)
 p1 = 0
 pl = 0
 puZ = 0
 For aktb = 1 To SL
  buch = Mid$(teststr, aktb, 1)
  If p1 = 0 Then
   If InStrB("0123456789", buch) <> 0 And buch <> "" Then p1 = aktb
  End If
  If p1 > 0 Then
   If buch = "." Then puZ = puZ + 1
   If InStrB("0123456789" + IIf(puZ > 2, vNS, "."), buch) = 0 Then
    pl = aktb - p1
    Exit For
   End If
  End If
 Next
 DatInStr = CDate(0)
 If p1 > 0 And pl > 0 Then
  ZwiStr = Mid$(teststr, p1, pl)
  If Not IsDate(ZwiStr) Then
   ZwiStr = ZwiStr + CStr(Jahr)
  End If
  If IsDate(ZwiStr) Then
   DatInStr = CDate(ZwiStr)
  End If
 End If
End Function ' DatInStr(TestStr$, Optional jahr) As Date

Function MedPlanAusAna(ByVal Pat_ID$)
 Dim rsNa As New ADODB.Recordset, i%, j%, inkl%, k%, dosstr$, dosier$()
 On Error GoTo fehler
' SET rsna = TabÖff("Anamnesebogen", "Pat_id")
 myFrag rsNa, "SELECT * FROM `anamnesebogen` WHERE pat_id = " & Pat_ID, adOpenStatic
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
      dosier(k) = REPLACE$(dosier(k), "€", "-")
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
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in MedPlanAusAna/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'MedPlanAusAna

' Wählt den ersten oder den aktuellen Medikamentenplan aus obakt = true = aktuellen, obakt = false = ersten
Function MedPlanNr&(ByVal Pat_ID$, ByVal obAkt, Optional ByVal VorDat As Date, Optional ByVal NurNr%, Optional ByVal nr&)
 Dim raFIM As New ADODB.Recordset
 On Error GoTo fehler
' sql1 = "SELECT " & IIf(obAkt, "max", "min") & "(mpnr) FROM `medplan` WHERE pat_id = " + CStr(Pat_id)
' IF vordat <> CDate(0) THEN sql1 = sql1 + " AND zeitpunkt < " & DatFor_k(vordat + 1)
' sql1 = sql1 + " ORDER BY zeitpunkt" + IIf(obAkt, " DESC", "") + ", Medikament, feldnr"
 'Set rFIM = Dtb.OpenRecordset(sql1, dbOpenDynaset)
 If nr <> 0 Then
  MedPlanNr = nr
 Else
'  sql1 = "SELECT " & IIf(obAkt, "max", "min") & "(mpnr) AS mpmax FROM `medplan` WHERE pat_id = " + CStr(Pat_id)
  sql1 = "SELECT " & IIf(obAkt, "max", "min") & "(mpnr) mpmax FROM medplan mp1 WHERE pat_id = " & Pat_ID & " AND zeitpunkt = (SELECT " & IIf(obAkt, "max", "min") & "(zeitpunkt) FROM medplan mp0 WHERE pat_id = mp1.pat_id)"
  If VorDat <> CDate(0) Then sql1 = sql1 + " AND zeitpunkt < " & DatFor_k(VorDat + 1)
'  raFIM.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
  myFrag raFIM, sql1
' IF NOT ISNULL(raFIM.Fields(0)) THEN
  If Not raFIM.BOF And Not IsNull(raFIM!mpmax) Then
   MedPlanNr = raFIM!mpmax
   If NurNr Then Exit Function
  Else
   Exit Function
  End If
 End If
 sql1 = "SELECT * FROM `medplan` WHERE pat_id = " & Pat_ID & " AND mpnr = " & MedPlanNr
 Set raFIM = Nothing
' raFIM.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
 myFrag raFIM, sql1
 ' END IF
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
     Dos(0, MedNr) = IIf(IsNull(raFIM!mo), vNS, raFIM!mo)
     Dos(1, MedNr) = IIf(IsNull(raFIM!mi), vNS, raFIM!mi)
     Dos(2, MedNr) = IIf(IsNull(raFIM!nm), vNS, raFIM!nm)
     Dos(3, MedNr) = IIf(IsNull(raFIM!ab), vNS, raFIM!ab)
     Dos(4, MedNr) = IIf(IsNull(raFIM!Zn), vNS, raFIM!Zn)
    End If
    raFIM.Move 1
   Loop
  End If ' not bof
' END IF ' NOT ISNULL
 Exit Function
fehler:
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in MedPlanNr/" + App.path)
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
 '   IF DosHäuf THEN Exit Function
   End If
  Next
 End If ' MedNr> -1
End Function 'DosHäuf%(MedNr%)

#If doppelt Then
' 20.3.17: folgende Funktion sollte jetzt durch SQL-Funktionen RRsyst und RRdiast erübrigt sein
' aufgerufen in rrParsen, und GetPrRR, do_RRParse
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
    T1 = left$(erg, pos - 1)
    If Len(T1) > 2 Then
     If Mid$(T1, Len(T1) - 2, 1) = "," And InStrB("0123456789", Mid$(T1, Len(T1) - 1, 1)) <> 0 Then T1 = left$(T1, Len(T1) - 2)
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
    T1a = left$(erg, 3)
    If IsNumeric(T1a) Then
     RRdiast = Val(T1a)
     erg = Mid$(erg, 4)
    Else
     T1a = left$(T1a, 2)
     If IsNumeric(T1a) Then
      RRdiast = Val(T1a)
      erg = Mid$(erg, 2)
     End If
    End If
    If InStr(erg, "(") > InStr(erg, "/") Then
     T1a = Mid$(erg, InStr(erg, "("))
     If InStr(T1a, ")") > 1 Then
      T1a = left$(T1a, InStr(T1a, ")") - 1)
      If IsDate(T1a) Then Zp = CDate(T1a)
     End If
    End If
   Else ' pos > 1
    pos = InStr(erg, "syst")
    If pos > 1 Then
     T1 = left$(erg, pos - 1)
     For i = Len(T1) To 0 Step -1
      If InStrB("0123456789", Mid$(T1, i, 1)) <> 0 And Mid$(T1, i, 1) <> "" Then Exit For
      T1 = left$(T1, Len(T1) - 1)
      If T1 = "" Then Exit For
     Next
     If Len(T1) > 2 Then
      If Mid$(T1, Len(T1) - 2, 1) = "," And InStrB("0123456789", Mid$(T1, Len(T1) - 1, 1)) <> 0 And Mid$(T1, Len(T1) - 1, 1) <> "" Then T1 = left$(T1, Len(T1) - 2)
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
       T1a = left$(T1a, Len(T1a) - 2)
       For i = Len(T1a) To 0 Step -1
        If T1a = vNS Then Exit For
        If InStrB("0123456789", Mid$(T1a, i, 1)) <> 0 And Mid$(T1a, i, 1) <> vNS Then Exit For
        T1 = left$(T1a, Len(T1a) - 1)
       Next
       If Len(T1a) > 2 Then
        If Mid$(T1a, Len(T1a) - 2, 1) = "," And InStrB("0123456789", Mid$(T1a, Len(T1a) - 1, 1)) <> 0 And Mid$(T1a, Len(T1a) - 1, 1) <> "" Then T1a = left$(T1a, Len(T1a) - 2)
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dodoRRParse/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dodoRRParse
#End If

' in DMPString
Function WieTabak(Pat_ID&) As ZigSt
  Dim raZig As New ADODB.Recordset
  Dim DiText$, DiSich$
  WieTabak = doTabakSt(Pat_ID)
  If Pat_ID <> 0 Then
'   Call raZig.Open("SELECT * FROM `diagnosen` WHERE pat_id = " & Pat_id & " AND icd LIKE '" & "F17" & "%" & "' AND diagsicherheit IN (""G"",""V"",""Z"")  AND COALESCE(Dggel,0)=0 ", DBCn, adOpenDynamic, adLockReadOnly)
   myFrag raZig, "SELECT DiagText,DiagSicherheit FROM diagview WHERE pat_id = " & Pat_ID & " AND gzicd LIKE 'F17%'"
   If Not raZig.EOF Then
    DiText = raZig!DiagText
    DiSich = raZig!DiagSicherheit
    If InStrB(DiText, "Z.n") <> 0 Or DiSich = "Z" Then
     If WieTabak <> früher And WieTabak <> vorlangem Then WieTabak = früher
    Else
     WieTabak = aktuell
    End If
   End If ' Not raZig.EOF Then
  End If ' Pat_id <> 0 Then
End Function ' WieTabak

Function testvergleicheT()
 Dim rs As New ADODB.Recordset, AusS$, e1 As ZigSt, E2 As ZigSt, e1W$, e2W$, Zahl&
 Call Lese.ProgStart
 Open pVerz & "tabakst.txt" For Output As #331
 myFrag rs, "SELECT d.icd, d.diagsicherheit, a.* FROM `anamnesebogen` a LEFT JOIN `diagnosen` d ON a.pat_id = d.pat_id AND d.icd LIKE 'F17%'"
 Do While Not rs.EOF
  e1 = doTabakSt(rs!Pat_ID)
  E2 = doTabakStAlt(rs!Tabak)
  If e1 = aktuell Then e1W = "Aktuell" Else If e1 = früher Then e1W = "Früher " Else If e1 = nie Then e1W = "Nie    " Else e1W = "VorLang"
  If E2 = aktuell Then e2W = "Aktuell" Else If E2 = früher Then e2W = "Früher " Else If E2 = nie Then e2W = "Nie    " Else e2W = "VorLang"
  If e1 <> E2 And Not (rs!tabakex = "j" And E2 = nie) And rs!Pat_ID <> 51 And rs!Pat_ID <> 602 And Not (e1 = früher And E2 = vorlangem And IsNumeric(rs!tabakbis)) And Not (e1 = nie And (E2 = früher Or E2 = aktuell)) And Not (e1 = vorlangem And E2 = früher) And Not (e1 = früher And E2 = vorlangem) Then
'  IF e1 = früher AND e2 = vorlangem AND IsNumeric(rs!tabakbis) THEN
'  IF e1 = früher AND e2 = vorlangem THEN
  If True Then
   AusS = e1W & " " & e2W & " " & Right$(Space(5) & rs!Pat_ID, 5) & " " & left$(rs!Nachname & Space(15), 15)
   AusS = AusS & left$(rs!ICD & rs!DiagSicherheit & Space(7), 7) & " "
   AusS = AusS & left$(rs!Tabak & Space(70), 70) & " "
   AusS = AusS & left$(rs!tabakakt & Space(15), 10) & " "
   AusS = AusS & left$(rs!tabakex & Space(15), 10) & " "
   AusS = AusS & left$(rs!tabakbis & Space(15), 15) & " "
   AusS = AusS & left$(rs!tabakmenge & Space(15), 10) & " "
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

 Function doTabakSt(Pat_ID&) As ZigSt
  Dim rs As New ADODB.Recordset, tStr$
  On Error GoTo fehler
  Call Lese.ProgStart
  myFrag rs, "SELECT vorgestellt, tabakex,tabakakt,tabakbis,tabakmenge FROM `anamnesebogen` WHERE pat_id = " & Pat_ID, adOpenStatic
  If Not rs.BOF Then
  If rs!tabakakt = "j" Or InStrB(rs!tabakakt, "gelegentlich") <> 0 Or InStrB(rs!tabakakt, "ja") <> 0 Then
   doTabakSt = aktuell
  ElseIf IsNull(rs!tabakex) Then
nie:
   doTabakSt = nie
  ElseIf rs!tabakex = "-" Or rs!tabakex = "" Or rs!tabakex = "n" Or rs!tabakex = "?" Or left$(rs!tabakex, 2) = "n " Or rs!tabakex = "nein" Or rs!tabakex = "Nein" Or rs!tabakex = "nie" Or rs!tabakex = "Nie" Or InStrB(rs!tabakex, "selber nicht") <> 0 Then
   GoTo nie
  Else
   doTabakSt = früher
   If Not IsNull(rs!tabakbis) Then
    tStr = LTrim$(rs!tabakbis)
    If left$(tStr, 3) = "bis" Then tStr = LTrim$(Mid$(tStr, 4))
    If left$(tStr, 3) = "ca." Then tStr = LTrim$(Mid$(tStr, 4))
    If InStrB(tStr, "-") <> 0 Then
     tStr = LTrim$(Mid$(tStr, InStr(tStr, "-") + 1))
    End If
    If IsDate(tStr) Then
     If Now() - CDate(tStr) > 12 * 365 Then
      doTabakSt = vorlangem
     End If
    Else
     If IsNumeric(left$(tStr, 4)) Then tStr = left$(tStr, 4)
     If IsNumeric(tStr) Then
      If Not IsNull(rs!Vorgestellt) Then
       If Year(rs!Vorgestellt) - tStr >= 12 Then
        doTabakSt = vorlangem
       End If
      End If
     ElseIf tStr Like "*vor *" Then
      tStr = Mid$(tStr, InStr(tStr, "vor ") + 4)
      If tStr Like "*J*" Then
       tStr = left$(tStr, InStr(tStr, "J") - 1)
      ElseIf tStr Like "*a*" Then
       tStr = left$(tStr, InStr(tStr, "a") - 1)
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
  End If ' not rs.bof
      Exit Function
fehler:
     Dim AnwPfad$
    #If VBA6 Then
     AnwPfad = CurrentDb.name
    #Else
     AnwPfad = App.path
    #End If
    Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doTabakSt/" + AnwPfad)
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
   If left$(Text, 1) = "j" Then
    sp = Split(Text, "Frühergeraucht?")
    If Trim$(sp(0)) = ",wieviel?" Or obNein(left$(sp(0), 1)) Then
     doTabakStAlt = nie
    Else
     If rSnpos > 0 Then
      If obNein(rSn) Then doTabakStAlt = früher Else If (obNein(sp(0)) Or InStrB(rSn, "j") <> 0) Then doTabakStAlt = aktuell
     Else
      If LCase$(left$(Text, 1)) = "j" Then
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
   ElseIf left$(rSn, 1) = "j" Then 'obNein(rSn) = 0 THEN 4.4.07 Pat. 1716
    doTabakStAlt = aktuell
   ElseIf left$(Text, 1) = "n" Or Text = "Frühergeraucht? n" Or Text = "Frühergeraucht? -" Or left$(Text, 1) = "," Then
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
      jahrk = left$(jahrk, InStr(jahrk, " ") - 1)
     End If
     If InStrB(jahrk, ".LJ") <> 0 Then
      jahrk = left$(jahrk, InStr(jahrk, ".LJ") - 1)
     End If
     If InStrB(jahrk, "/") <> 0 Then
      jahrk = Mid$(jahrk, InStr(jahrk, "/") + 1)
     End If
     If IsDate(jahrk) Then
      Datu = CDate(jahrk)
      obdatum = -1
     Else
      jahrk = left$(jahrk, 4)
     If IsNumeric(jahrk) Then
      obdatum = -1
     Else
      If IsNumeric(left$(jahrk, 2)) Then
       jahrk = left$(jahrk, 2)
       obdatum = -1
      End If
     End If
    End If
    
   End If ' MinusStelle = 1 OR bisStelle = 1 THEN
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
'        IF Not obdatum OR (obdatum AND datu > Now + 15 * 365) THEN
' muß noch getestet werden
         doTabakStAlt = früher
       Else
        doTabakStAlt = aktuell
       End If
      End If
     End If
    End If
   End If
  End If ' NOT ISNULL(Text
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doTabakStAlt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doTabakStAlt

Function sonz$(Quelle$)
 Dim i&, ZS$, bz$, bqb$, bq%
 On Error GoTo fehler
 For i = 1 To Len(Quelle)
  bqb = Mid$(Quelle, i, 1)
'  bq = AscW(bqb)
  Select Case bqb
   Case "&": bz = "&amp;"
   Case """": bz = "&quot;"
   Case "'": bz = "&#039;"
   Case "<": bz = "&lt;"
   Case ">": bz = "&gt;"
   Case Else: bz = bqb
  End Select
  sonz = sonz + bz
 Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZSUh/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' sonz
 
Public Function pttest$(ByRef quel$)
  Dim Ziel As String * 64000
  Dim pq&, pz&
  pq = VarPtr(quel)
  pz = VarPtr(Ziel)
  Debug.Print pq, pz
  Call CopyMem(Ziel, "a", 1)
  Debug.Print pq, pz
  Debug.Print pz, Ziel
End Function ' pttest

 
' in tuBriefStandalone, einzeintr, Epikrise; zu koordinieren mit gleichnamiger Mariadb-Funktion
Public Function zuh$(ByRef Quelle$, Optional mitabsatz%, Optional obgross%)
 Dim i&, bqb$, bq%, zul13%
' zuh = ""
 For i = 1 To Len(Quelle)
  bqb = Mid$(Quelle, i, 1)
  bq = Asc(bqb)
  If bq = 10 And zul13 Then GoTo weiter
  zul13 = 0
  Select Case bq
   Case 13
    If i < Len(Quelle) Then
'     zuh = zuh & "</w:t><w:br/><w:t>" ' das letzte Zeilenumbruchszeichen wird durch Absatzformatierung erfüllt
'     zuh = zuh & "</w:t>" & IIf(mitabs, "</w:r></w:p><w:p><w:pPr><w:tabs><w:tab w:val=""left"" w:pos=""1134""/></w:tabs><w:ind w:left=""1134"" w:hanging=""1134""/></w:pPr><w:r><w:rPr><w:rStyle w:val=""s" & IIf(obgross, "24", "18") & "s""/></w:rPr><w:tab/>", "<w:br/>") & "<w:t>" ' das letzte Zeilenumbruchszeichen wird durch Absatzformatierung erfüllt
     zuh = zuh & "</w:t>" & IIf(mitabsatz, "</w:r></w:p><w:p><w:pPr><w:pStyle w:val=""hang""/></w:pPr><w:r><w:rPr><w:rStyle w:val=""s" & IIf(obgross, "24", "18") & "s""/></w:rPr>", "<w:br/>") & "<w:t>" ' das letzte Zeilenumbruchszeichen wird durch Absatzformatierung erfüllt
    End If
    zul13 = True
   Case 38, 60, 62, Is >= 128
    zuh = zuh & "&#" & bq & ";"
   Case Else
    zuh = zuh & bqb
  End Select ' bq
weiter:
 Next i
End Function ' zuh

' in tubriefStandalone, do_DMPAusgebStandAlone
Function zsuh$(ByRef Quelle$)
 Dim i&, ZS$, bz$, bqb$, bq%
 On Error GoTo fehler
 For i = 1 To Len(Quelle)
  bqb = Mid$(Quelle, i, 1)
'  If bqb = "ü" Then Stop
  bq = AscW(bqb)
  Select Case bq
   Case Is < 128
      bz = bqb
   Case 128: bz = "â‚¬" ' Chr(&HE2) & Chr(&H82) & Chr(&HAC) ' €
   Case 129: bz = Chr$(&HC2) & Chr$(&H81)
   Case 130: bz = "â€š" ' ‚
   Case 131: bz = "Æ’"  ' ƒ
   Case 132: bz = "â€ž" ' „
   Case 133: bz = "â€†" ' …
   Case 134: bz = "â€ " '
   Case 135: bz = "â€¡"
   Case 136: bz = "Ë†"
   Case 137: bz = "â€°"
   Case 138: bz = "Å "
   Case 139: bz = "â€¹"
   Case 140: bz = "Å’"
   Case 141: bz = Chr$(&HC2) & Chr$(&H8D)
   Case 142: bz = "Å½"
   Case 143: bz = Chr$(&HC2) & Chr$(&H8F)
   Case 144: bz = Chr$(&HC2) & Chr$(&H90)
   Case 145: bz = Chr$(&HE2) & Chr$(&H80) & Chr$(&H98)
   Case 146: bz = "â€™"
   Case 147: bz = "â€œ"
   Case 148: bz = Chr$(&HE2) & Chr$(&H80) & Chr$(&H9D)
   Case 149: bz = "â€¢"
   Case 150: bz = "â€“"
   Case 151: bz = "â€”"
   Case 152: bz = "Ëœ"
   Case 153: bz = "â„¢"
   Case 154: bz = "Å¡"
   Case 155: bz = "â€º"
   Case 156: bz = "Å“"
   Case 157: bz = Chr$(&HC2) & Chr$(&H9D)
   Case 158: bz = "Å¾"
   Case 159: bz = "Å¸"
   Case 8226:
        bz = "â€¢" ' •
   Case Is < 192
      bz = Chr$(&HC2) & bqb
   Case Is < 256
     bz = Chr$(&HC3) & Chr$(bq - 64)
   Case Is < 320
     bz = Chr$(&HC4) & Chr$(bq - 64)
   Case Is < 384
     bz = Chr$(&HC5) & Chr$(bq - 192)
   Case Else
     bz = "?"
  End Select
  zsuh = zsuh + bz
 Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZSUh/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' zsuh

#If zsusicher Then
Function zsuh_$(Quelle$)
 Dim i&, ZS$, bq$, bz$
 For i = 1 To Len(Quelle)
  bq = Mid$(Quelle, i, 1)
  Select Case bq
   Case "ä": bz = "Ã¤"
   Case "ö": bz = "Ã¶"
   Case "ü": bz = "Ã¼"
   Case "Ä": bz = "Ã„"
   Case "Ö": bz = "Ã–"
   Case "Ü": bz = "Ãœ"
   Case "ß": bz = "ÃŸ"
   Case "é": bz = "Ã©"
   Case "è": bz = "Ã¨"
   Case "à": bz = "Ã "
   Case "ù": bz = "Ã¹"
   Case "ò": bz = "Ã²"
   
   Case "ê": bz = "Ãª"
   Case "ô": bz = "Ã´"
   Case "û": bz = "Ã»"
   Case "î": bz = "Ã®"
   Case "í": bz = "Ã­"
   Case "ì": bz = "Ã¬"
   Case "Á": bz = "Ã"
   Case "É": bz = "Ã‰"
   Case "Í": bz = "Ã"
   Case "Ó": bz = "Ã“"
   Case "Ú": bz = "Ãš"
   Case "À": bz = "Ã€"
   Case "È": bz = "Ãˆ"
   Case "Ì": bz = "ÃŒ"
   Case "Ò": bz = "Ã’"
   Case "Ù": bz = "Ã™"
   Case Is < 128
    bz = bq
   Case Else
     bz = bq
  End Select
  zsuh_ = zsuh_ + bz
 Next i
 Exit Function
End Function ' zsuh__
#End If

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
' Debug.Print "ZSU: " + ZSU + " -> " + Quelle
End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZSU/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ZSU

' in waehleinMO(Pat_id&), PatAuswahl.inTM_Click, AnBog.inTM_Click, Medarten.inTM_Click
Public Sub inMOAnz(Pat_ID&)
 Static fehlerangezeigt%
On Error GoTo fehler
'Clipboard.Clear
'Clipboard.SetText CStr(Pat_ID)
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
     pstr = Pat_ID 'adoRS!Pat_id
     For i = 1 To Len(pstr)
      Call cAutoit.send(" ")
      Call cAutoit.send(Mid$(pstr, i, 1))
    ' Call cAutoit.sleep(100)
     Next i
'     Debug.Print "adors!pat_id: " & pstr
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
    Dim hnd&, j&, k&, i%
    Const MOZ$ = "Medical Office - Zentrale"
    On Error GoTo fehler
    If Lese.MOBetr <> 0 Then
     Debug.Print "Lade '" & Pat_ID & "'"
     For i = 1 To 5
        hnd = FensterHandle(MOZ)
        If hnd <> 0 Then Exit For
        If i < 5 Then
         On Error Resume Next
          Select Case i
           Case 1: Shell "c:\medoff\medoff.exe"
           Case 2: Shell "c:\indamed\medoff.exe"
           Case 3: Shell "d:\medoff\medoff.exe"
           Case 4: Shell "d:\indamed\medoff.exe"
          End Select
          Pause 2000
         On Error GoTo fehler
        End If ' hnd = 0 And i < 5 Then
     Next i
     If hnd = 0 Then
      MsgBox "Aktivierung von MO nicht möglich, hnd=0"
     Else
      Pausenlänge = 10
      On Error GoTo dgfehler
      AppActivate "Diagnoseerfassung"
      Pause (Pausenlänge)
'      SendKeysEx "+{TAB}"
'      SendKeysEx "{RETURN}"
'      SendKeysEx "%{F4}"
      SendKeysEx "{ESC}"
      Pause (Pausenlänge)
      SendKeysEx "%{TAB}"
'      On Error Resume Next
'      AppActivate MOZ
'      Pause (100 * Pausenlänge)
zumo:
      On Error GoTo mofehler
      AppActivate MOZ, True
      If Err.Number <> 0 Then Exit Sub
      Debug.Print "nach AppActivate"
      Screen.MousePointer = vbDefault
      Debug.Print "screen.width:", Screen.Width, "screen.height:", Screen.Height
      For j = 200 To 200 ' 6000 To 6000 Step -100
       For k = 20 To 20 ' 600 To 600 Step -10
        Debug.Print "j:", j, "k:", k
        Mausklick vbLeftButton, j, k ' Screen.Width / 2, Screen.Height / 2
'        Mausklick vbLeftButton, j, k ' Screen.Width / 2, Screen.Height / 2 ' auskommentiert 23.6.25
'        MsgBox "j: " & j & " k: " & k
'        AppActivate "Medical Office - Zentrale", True
        Pause (Pausenlänge)
       Next k
      Next j
      DoEvents
     ' SendK VK_F4
      Debug.Print "vor Sendkeys F4"
'      Sendkeys "{F4}", False
       SendKeysEx "{F4}"
'      Dim WshShell
'      Set WshShell = CreateObject("WScript.Shell")
'      WshShell.RunOnclient = True
'      WshShell.Sendkeys "{F4}"
'      Pause (Pausenlänge)
'      Sendkeys Pat_id, False
      SendKeysEx Pat_ID
'     For i = 1 To Len(CStr(Pat_ID))
'      SendK Mid(CStr(Pat_ID), i, 1)
'     Next i
      Pause (Pausenlänge)
      DoEvents
'     SendK vk_return
'      Sendkeys "{ENTER}", False
      SendKeysEx "{ENTER}"
      DoEvents
      Pause (Pausenlänge)
'     SendK vk_return '
'      Sendkeys "{ENTER}", False
      SendKeysEx "{ENTER}"
     End If
   Else ' Lese.MoBetr => Turbomed
     hnd = FensterHandle("TurboMed.exe")
     If hnd <> 0 Then
      On Error Resume Next
      AppActivate "TURBOMED", True
      If Err.Number <> 0 Then Exit Sub
      Pause (Pausenlänge)
      Sendkeys "{ESC}", True
      Pause (Pausenlänge)
      Sendkeys "{ESC}", True
      Pause (Pausenlänge)
      Sendkeys "{ESC}", True
      Pause (Pausenlänge)
      Sendkeys "{ENTER}", True
      Pause (Pausenlänge)
      Sendkeys "{F12}", True
      Pause (Pausenlänge)
      Sendkeys "p", True
      Pause (Pausenlänge)
      Sendkeys "{bs}" & Pat_ID & "", True
      Pause (Pausenlänge)
      Sendkeys "{ENTER}", True
      Pause (Pausenlänge)
      Sendkeys "{F3}", True
      Pause (Pausenlänge)
      Sendkeys "%{F1}", True
      On Error GoTo fehler
    '  Pause (Pausenlänge)
    '  AppActivate "TurboMed", True
     End If
   End If ' Lese.MoBetr Else
  Exit Sub
dgfehler:
 Resume zumo
mofehler:
' Select Case MsgBox("Fehler beim Aufruf von " & Pat_ID & " in Medical Office. Bitte diese Anwendung überprüfen!", vbAbortRetryIgnore, "Aufgefangener Fehler in waehleinMO/" + App.path)
'  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
'  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
'  Case vbIgnore: Call MsgBox("Setze fort"): Exit Sub
' End Select
 Debug.Print Err.Number, Err.Description
 If Err.Number = 70 And Not fehlerangezeigt Then
  MsgBox "Kann Pat. in MO nicht aufrufen, '" & Err.Description & "', ggf. dieses Programm als Administrator aufrufen"
  fehlerangezeigt = True
 End If ' Err.Number = 50 And Not fehlerangezeigt Then
 Exit Sub
#End If
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("Fehler beim Aufruf von Medical Office. Bitte diese Anwendung überprüfen!", vbAbortRetryIgnore, "Aufgefangener Fehler in waehleinMO/" + App.path)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' inMOAnz(Pat_id&)

Public Function GetRandomNum&(MIN&, MAX&)
    GetRandomNum = Int(Rnd * (MAX - MIN + 1) + MIN)
End Function ' GetRandomNum

'Public FUNCTION s2test(Pat_id&)
'Lese.ProgStart
'Dim z&
'Dim r2 As New ADODB.Recordset
'   myFrag r2, "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_id & " AND abkü RLIKE '^hba[1c]' AND CAST(wert AS decimal) < 22 UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_id & " AND abkü RLIKE '^hba[1c]' AND CAST(wert AS decimal) < 22) i GROUP BY pat_id,zeitpunkt,abkü,wert ORDER BY zeitpunkt " '
'   Do While Not r2.EOF
'    z = z + 1
'    r2.MoveNext
'   Loop
'   Debug.Print z
'   z = 100
'   SET r2 = hollabor(Pat_id, "HBA[1C]", 0, 22, 0, z)
'Debug.Print z
'End Function

' zahl liefert die Zahl der Datensätze zurück, wenn nicht -1 übergeben wird
Public Function hollabor(Optional PatID& = 0, Optional Abkü$ = "", Optional zpkl As Date = 0, Optional wertkl& = 0, Optional obnachgruppe% = 0, Optional ByRef Zahl& = -1, Optional Einheit$ = "", Optional gz& = -1, Optional dzz& = -1, Optional dSL As SortierListe, Optional obUpdate%, Optional obneu%, Optional ohnebr%) As ADODB.Recordset
 Static rs As New ADODB.Recordset, sql$
 Static altPID&, altab$, altzpkl#, altwk&, altobng%, alteh$, altgz&, altdz&
 Static CursorTp As ADODB.CursorTypeEnum, LockTp As ADODB.LockTypeEnum
 Static hdSL As New SortierListe ' dzz = datumszahl, wenn nicht nach Gruppen gruppiert (also nach Zeitpunkt sortiert)
#If gehtnicht Then
 If obneu Then
  Set rs = Nothing: Set hdSL = Nothing: sql = "": altPID = 0: altab = "": altzpkl = 0: altwk = 0: altobng = 0: alteh = "": altgz = 0: altdz = 0
  CursorTp = 0: LockTp = 0
 End If ' obneu
#End If
 If obneu Then altPID = -1
 Dim SD As SortierDatum
 Dim vgru$ ' vorherige Gruppe
 On Error GoTo fehler
 If CursorTp = adOpenDynamic And PatID = altPID And altab = Abkü And altzpkl = zpkl And altwk = wertkl And altobng = obnachgruppe And alteh = Einheit Then
  If rs.State = 0 Then
   sql = rs.Source
   Set rs = Nothing
   rs.Open sql, DBCn, adOpenStatic, adLockReadOnly
  End If
  If Not rs.BOF Then rs.MoveFirst
 Else  ' CursorTp = adOpenDynamic And PatID = altPID And altab = Abkü And altzpkl = zpkl And altwk = wertkl And altobng = obnachgruppe And alteh = Einheit Then
  If PatID <> altPID Then
   altgz = 0
   altdz = 0
   Set hdSL = Nothing
  End If
  Set rs = Nothing
'#Const problematisch = True
#If problematisch Then
 If obUpdate Then
#End If
  CursorTp = adOpenDynamic
  LockTp = adLockOptimistic
#If problematisch Then
 Else
  CursorTp = adOpenUnspecified
  LockTp = adLockReadOnly
 End If
#End If
#If problematisch Then
  'myEFrag "SET @patid=" & CStr(PatID)
  'myEFrag "SET @abkue='" & Abkü & "'"
  'myEFrag "SET @einheit='" & Einheit & "'"
  'sql = "SET @zpkl="
  sql = ""
  If zpkl = 0 Then sql = sql + "0" Else sql = sql + DatFor_k(zpkl)
  'myEFrag sql
  myEFrag "SET @patid=" & CStr(PatID) & ",@abkue='" & Abkü & "',@einheit='" & Einheit & "',@zpkl=" & sql & ",@wertkl=" & wertkl
' rs.CursorLocation = adUseServer
  sql = "SELECT * FROM geslab" + IIf(obnachgruppe, " WHERE (reihe <> 999 OR ISNULL(reihe)) GROUP BY gruppe, reihe, abkü, einheit,ung,ong ORDER BY gruppe,reihe", " ORDER BY zeitpunkt DESC")
'  rs.Open sql, DBCn, adOpenDynamic, adLockReadOnly
  myFrag rs, sql, CursorTp, DBCn, LockTp
#Else
 Dim par$
 ' 4.7.20: wert statt einheit eingesetzt, da in labor1 und labor2 offenbar verschiedene Einheiten verwendet werden, z.B. ml/min = ml/mn/1.73 m²
' par = IIf(obnachgruppe, " WHERE (reihe <> 999 OR ISNULL(reihe)) GROUP BY gruppe, reihe, abkü, einheit,ung,ong)i ORDER BY gruppe,reihe", IIf(Abkü <> "", "WHERE abkü=""" & Abkü & """" & IIf(Einheit <> "", " AND einheit =""" & Einheit & """", "") & IIf(zpkl <> 0, " AND zeitpunkt<" & Format(zpkl, "yyyymmdd"), ""), "") & " GROUP BY zeitpunkt DESC,abkü,wt)i ORDER BY zeitpunkt DESC")
 par = IIf(obnachgruppe, " GROUP BY gruppe,reihe,abkü,einheit)i ORDER BY gruppe,reihe", IIf(Abkü <> "", "AND abkü=""" & Abkü & """" & IIf(Einheit <> "", " AND einheit =""" & Einheit & """", "") & IIf(zpkl <> 0, " AND zeitpunkt<" & Format(zpkl, "yyyymmdd"), ""), "") & " GROUP BY zeitpunkt DESC,abkü,wt)i ORDER BY zeitpunkt DESC")

' myEFrag("flush tables")
'  rs.Open "call geslabdp(" & CStr(PatID) & ",'" & par & "')", DBCn, adOpenStatic, adLockReadOnly
' myFrag rs, "CALL geslabdp" & IIf(ohnebr, "ohnebr", "") & "(" & CStr(PatID) & ",'" & par & "')", CursorTp, DBCn, LockTp
 myFrag rs, "CALL geslabneu(" & CStr(PatID) & ",'" & IIf(ohnebr, "", "<br>") & "','" & par & "')", CursorTp, DBCn, LockTp
' Dim satzzahl&
 If Not rs.BOF Then
'  Do While Not rs.EOF
'   satzzahl = satzzahl + 1
''   Debug.Print "Satzzahl: " & satzzahl
'   rs.MoveNext
'  Loop
'  Debug.Print "Satzzahl: " & satzzahl
'  rs.MoveFirst
'  Do While Not rs.EOF
'   Debug.Print rs!Zeitpunkt, rs!Abkü, rs!Wert, rs!Einheit
'   rs.MoveNext
'  Loop
  rs.MoveFirst
 End If ' Not rs.BOF Then
#End If
 End If ' CursorTp = adOpenDynamic And PatID = altPID And altab = Abkü And altzpkl = zpkl And altwk = wertkl And altobng = obnachgruppe And alteh = Einheit Then
#If langsamer Then
 If Zahl <> -1 Then Zahl = myEFrag("SELECT COUNT(0) Zahl FROM (" & sql & ")")!Zahl
#Else
 If Zahl <> -1 Then If rs.BOF Then Zahl = 0 Else Zahl = rs!dszahl
#End If
' das folgende geht bei ca. 1500 Datensätzen in 0,05s statt mit COUNT(0) in 0,15s
 If gz <> -1 Or dzz <> -1 Then ' wenn Zahl gewünscht wird
  If (gz <> -1 And altgz = 0) Or (dzz <> -1 And altdz = 0 And Not obnachgruppe) Then ' ... aber nicht vorliegt
   altgz = 0
   If altPID <> Pat_ID Or Not obnachgruppe Then
    altdz = 0
    Set hdSL = Nothing
   End If
   Do While Not rs.EOF ' ... dann ausrechnen
    If rs!Gruppe <> vgru Then
     altgz = altgz + 1
     vgru = rs!Gruppe
    End If
    ' wenn der Patient einmal ohne obnachgruppe und dann mit aufgerufen wird, dann kann wg. Performance altdz vom vorigen Aufruf jetzt zurückgeliefert werden
    If altPID <> Pat_ID Or Not obnachgruppe Then
     Set SD = New SortierDatum
     SD.Datum = Int(rs!Zeitpunkt)
     If hdSL.GetItem(SD) Is Nothing Then
      hdSL.sCAdd SD
      altdz = altdz + 1
     End If
    End If ' not obnachgruppe
    rs.MoveNext
   Loop
   If Not rs.BOF Then rs.MoveFirst
  End If ' ... dann
  If gz <> -1 Then gz = altgz
  If dzz <> -1 Then
   dzz = altdz
   Set dSL = hdSL
  End If
 End If ' gz <> -1 Or dzz <> -1 Then ' wenn Zahl gewünscht wird
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), vNS, CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in hollabor/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' hollabor

Public Function testlab(pid&)
 Dim rs As New ADODB.Recordset
 Lese.ProgStart
' myEFrag ("SET @patid=" & pid)
' myFrag rs, "SELECT * FROM geslab"
Const obnachgruppe% = 1
Dim par$
'par = IIf(obnachgruppe, " WHERE (reihe <> 999 OR ISNULL(reihe)) GROUP BY gruppe, reihe, abkü, einheit,ung,ong)i ORDER BY gruppe,reihe", " GROUP BY zeitpunkt DESC,abkü,einheit)i ORDER BY zeitpunkt DESC")
par = IIf(obnachgruppe, " GROUP BY gruppe,reihe,abkü,einheit,ung,ong)i ORDER BY gruppe,reihe", " GROUP BY zeitpunkt DESC,abkü,einheit)i ORDER BY zeitpunkt DESC")
'myFrag rs, "CALL geslabdp(" & pid & ",'" & par & "')"
myFrag rs, "CALL geslabneu(" & pid & ",'<br>','" & par & "')"
 TabAusgeb rs, Lese, True, , , , , , "Labor " & pid
End Function ' testlab(pid&)

'#If False THEN
Public Function testhl()
 Dim pid&
 pid = 68239
' für HAAkt z.B. geht das nicht, da es kein Lese enthält
 Call Lese.ProgStart
 Dim rs As ADODB.Recordset
 Set rs = hollabor(pid, "")
' für HAAkt z.B. geht das nicht, da es kein Lese enthält
 TabAusgeb rs, Lese, True, , , , , , "Labor " & pid, , , "Labor " & pid & " " & myEFrag("select gesnameg(" & CStr(pid) & ")").Fields(0)
 Debug.Print rs.Source
 rs.Close
' SET rs = hollabor(2, "Ery")
' TabAusgeb rs, Lese, True, , , , , , "Labor 2"
End Function
'#END IF

Public Function zuZahl$(Str$)
 Dim i&, erg$, buch$
 For i = 1 To Len(Str)
  buch = Mid$(Str, i, 1)
  If buch Like "[0-9,.]" Then
   If buch = "." Then buch = ","
   erg = erg & buch
  Else
   If erg <> "" Then Exit For
  End If
 Next i
 zuZahl = erg
End Function ' zuZahl(str$)

#If nurkurz Then
Public Function dakumw()
 Dim rNa As ADODB.Recordset, rEi As ADODB.Recordset, rAf&, sql$, ErrNr&, ErrDes$
 Lese.ProgStart
 sql = "select pat_id,IF(DAYNAME(REPLACE(REPLACE(MID(Inhalt, 30, INSTR(Inhalt,',')-30),',','.'),':','.'))IS NULL,18991230,STR_TO_DATE(REPLACE(REPLACE(MID(Inhalt, 30, InStr(Inhalt, ',') - 30),',','.'),':','.'),'%d.%m.%y'))dakab,inhalt from eintraege where art='dak' order by pat_id, zeitpunkt"
 myFrag rEi, sql, adOpenStatic, DBCn, adLockReadOnly, , rAf, , ErrNr, ErrDes
 If ErrNr <> 0 Then Stop
 If rEi!dakab = "" Then Stop
 Do While Not rEi.EOF
  sql = "update namen set dakab='" & rEi!dakab & "' where pat_id=" & rEi!Pat_ID
  myFrag rNa, sql, adOpenStatic, DBCn, adLockReadOnly, , rAf, , ErrNr, ErrDes
  If ErrNr <> 0 Then Stop
  Debug.Print sql, rAf
  rEi.MoveNext
 Loop
End Function
#End If
