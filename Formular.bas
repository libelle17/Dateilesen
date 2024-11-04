Attribute VB_Name = "Formular"
Option Explicit
Dim ag As New CString
Dim donr$
Dim bmnr&
Public Const KVNr = "6419153"
Public Const BSNR = KVNr & "00"

Const HADBName$ = "haerzte"
Type hatyp
 kvnu As String
 Vorname As String
 Nachname As String
' beme AS String
' gemmit AS String
' arzttyp AS String
' ort AS String
' tel1 AS String
' fax1 AS String
' email AS String
' anrede AS String
' titel AS String
' straße AS String
' plz AS String
' aktzeit As Date
' dmpt1 As Date
' dmpt2 As Date
End Type
'Const DMPVorlage$ = uVerz & "DMP-Vorlage.dot"
'Const FaxSendDatei$ = uVerz & "test1.txt"
'Declare FUNCTION Tüt& Lib "kernel32" Alias "Beep" (ByVal dwFreq AS Long, ByVal dwDuration AS Long)
Const übertragen$ = "übertragen"
Public Enum DSiTyp
 gesi& = 0
 Va&
 Zn&
 AuS&
End Enum
Dim NKrStr$(), Nflag$(), DSi() As DSiTyp ' für KRAdd
#Const mitab = True ' auch noch in Lese
#If mitab Then
'Public Const CStrMy$ = "DRIVER={MySQL ODBC 5.1 Driver};server=" & LiName & ";user=...;pwd=...;database="
Private CStrMy$
#End If
#Const debu = 0 ' noch in Importiert
#If debu <> 0 Then
       Public T0!, Takt!, Tvor!, dnr&
#End If
'Public Const wdLineStyleSingle% = 1
'Public Const wdWindowStateMaximize% = 1
'Public Const wdFindContinue% = 1
'Public Const wdAlignTabLeft% = 0
'Public Const wdTabLeaderDots% = 1
'Public Const wdReplaceAll% = 2
'Public Const wdWord9TableBehavior% = 1
'Public Const wdAutoFitContent% = 1
'Public Const wdColorRed% = 255
'Public Const wdBorderLeft% = -2
'Public Const wdLineWidth050pt% = 4
'Public Const wdColorAutomatic& = -16777216
'Public Const wdBorderRight% = -4
'Public Const wdBorderTop% = -1
'Public Const wdBorderBottom% = -3
'Public Const wdBorderHorizontal% = -5
'Public Const wdBorderVertical% = -6
'Public Const wdLineStyleNone% = 0
'Public Const wdBorderDiagonalDown% = -7
'Public Const wdBorderDiagonalUp% = -8
'Public FNam$() ' FreigabeNamen
'Public FInh$() ' FreigabeInhalt
'Public FZ%     ' Freigabezahl
Public Const MarkierFarbe$ = 9803263 ' rötlich
Public DatPfad$, DatPfadPI$, RDatPfad$
Public HANrBf$() ' Aktuell für den Brief verwendeter 1. und 2. Hausarzt
Public QMdbAkt$
'Public rafDE As Adodb.Recordset ' für Diagnosen Export
Public raLau As New ADODB.Recordset ' für Anamnesebogen-Form_Current und Anamnesebogen-Form_Current2
'Dim aNaüm As New ADODB.Recordset ' Anamnesebogen in form_current
'Public TMInterface AS Object
'Public PcDokPfad$
Dim An1Pfad$, An2Pfad$, AnAPfad$, CheckPfad$, VorlagenPfad$
Dim Alter%, Pat_ID&, GebDat As Date, fm As Form, obWeib As Boolean
Dim DSNr%, altDS(), altpat_id
Public MDIICD$(30) ' ICD-Nummern im Feld MDI
Public MDIDiag$(30) ' Diagnosentext im Feld MDI
Public MDIn% ' Fallzahl für die Zeile in MDI
' Dim i%
Public CGr%
'Dim nzl ' Neue Zeile -> vbcrlf
Dim nzw$ ' neue Zeile in Winword
Dim lTherZP As Date ' Zeitpunkt des letzten Therapieplans
Public obkNeph% ' ob keine Nephropathie zu bestehen scheint (-1 => vorliegende Diagnose wird ignoriert)
Public obMakroAlb% ' ob eine Makroalbuminurie vorliegt
' Aus Epikrise zur allgemeinen Verfügung:
    Public Const FEZ% = 27, maxIcd% = 5
    Public Const flNep% = 2
    Public Const flNiI% = 3
    Public Const flAVK% = 7
    Public Const flHyp% = 14
    Public Const flDFS% = 15
    Public Const flPNP% = 16
    Public Const flLiph% = 21
    Public Const flChar% = 23
    Public Const flUlc% = 26
    Public Const flMFF% = 27 ' höhergradiges Fußsyndrom
    Public flag%(FEZ)
    Public DiagSi%(FEZ)
    Public DiagNamen$(FEZ)
    Public ic$(FEZ, maxIcd)
    Public un%(FEZ) ' Unterdiagnosen
' Begleiterkrankungen
    Public Const BEZ% = 4, bmaxIcd% = 4
    Public bflag%(BEZ)
    Public bDiagsi$(BEZ)
    Public bDN$(BEZ)
    Public bic$(BEZ, bmaxIcd - 1)
    Public bun%(BEZ) ' Unterdiagnosen
    Public Const bflHyt% = 1
    Public Const bflRau% = 4
'Public Const KVÄDatei$ = aVerz + "\KV-Ärzte.mdb" ' uverz & "Anamnese"
'Public KVÄDatei$
Public labPos% ' Laborwert positiv: -1 = positiv, 0 = negativ, 1 = V.a. positiv
'Public fpos&
'Declare FUNCTION RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey&, ByVal lpSubKey$, ByVal ulOptions&, ByVal samDesired&, phkResult&)
'Declare FUNCTION RegQueryValueEx& Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey&, ByVal lpValueName$, ByVal lpReserved&, lpType&, lpData AS Any, lpcbData&)         ' Note that IF you declare the lpData parameter$, you must pass it By Value.
'Declare FUNCTION RegCloseKey& Lib "advapi32.dll" (ByVal hKey&) ' unfertig, sollte ein Formular automatisch machen
'Declare FUNCTION RegEnumKey& Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey&, ByVal dwIndex&, ByVal lpName$, ByVal cbName&)
'Declare FUNCTION RegEnumValue& Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey&, ByVal dwIndex&, ByVal lpValueName$, lpcbValueName&, ByVal lpReserved&, lpType&, ByVal lpData$, lpcbData&)
' Deklarieren der nötigen API-Routinen:
' ... für das Excel-Beispiel
'Declare FUNCTION FindWindow Lib "user32" Alias _
'"FindWindowA" (ByVal lpClassName$, _
'                    ByVal lpWindowName&)&
Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
' für das Winword-Beispiel
Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
                    
'Store these declarations in the form wh
' ich is making the calls to Word
'Public Wapp AS Object 'Word.Application
Dim WordWasNotRunning As Boolean ' Flag For final word unload
'The main sub which detects word, AND cr
' eates the object, OR starts a new Word
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal SleepTime&)
Public Declare Function sndPlaySound32& Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName$, ByVal uFlags&)
'Public WinDir$
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096
Declare Function OpenClipboard& Lib "user32" (ByVal hwnd&)
Declare Function CloseClipboard& Lib "user32" ()
Declare Function GetClipboardData& Lib "user32" (ByVal wFormat&)
Declare Function GlobalLock& Lib "kernel32" (ByVal hMem&)
Declare Function GlobalUnlock& Lib "kernel32" (ByVal hMem&)
Declare Function lstrcpy& Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any)
Declare Function GetWindowsDirectory& Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer$, ByVal nSize&)
Declare Function RegEnumKey& Lib "advapi32.dll" _
                 Alias "RegEnumKeyA" ( _
                 ByVal hKey&, _
                 ByVal dwIndex&, _
                 ByVal lpName$, _
                 ByVal cbName&)
'Public rest AS DAO.Recordset
Dim Tabl ' Labortabelle, Funktionsübergreifend
Dim UStumm% ' Knöpfe neben den Diagnosen in Position setzen
Public Const laborAbfr$ = "SELECT n.Pat_ID AS Pat_ID,n.ZeitPunkt AS ZeitPunkt,n.FertigStGrad AS FertigStGrad,n.Abkü AS Abkü,l.Langtext AS Langtext,n.Wert AS Wert,n.Einheit AS Einheit,k.Kommentar AS Kommentar,n.AbsPos AS AbsPos,n.AktZeit AS AktZeit FROM (`laborlangtext` l INNER JOIN (laborkommentar k INNER JOIN `laborneu` n ON ((k.KommentarVW = n.KommentarVW))) ON ((l.LangtextVW = n.LangtextVW)))"

Dim psql$(10)

' gemeinsame Algorithmusdefinition zum Therapieartenermitteln, in doViewserstellen und TheraErmitt
Function therinit()
 psql(0) = "SET SESSION GROUP_CONCAT_MAX_LEN=15000;"
' psql(1) = "IF inpid IN('','0') THEN DELETE FROM therarten; ELSE DELETE FROM therarten WHERE FIND_IN_SET(pat_id,inpid)>0; END IF;"
' psql(1) = "IF inpids IN('','0') THEN DELETE FROM therarten; ELSE DELETE FROM therarten WHERE pat_id IN (inpid); END IF;"
 psql(1) = "DELETE FROM therarten WHERE (inpid='0' OR pat_id IN (inpid));"
 psql(2) = "SET @vzahl = ROW_COUNT(); "
' psql(3) = "INSERT INTO therarten(pat_id,zp,mpnr,therart,insart,grund,abspos,aktzeit,stbyte) " & vbCrLf & _
"SELECT pid,zp,mpnr,thart,ia,gru,abspos,NOW(),stbyte FROM ( " & vbCrLf & _
"WITH dsort AS ( -- da zum Vergleich mit dem Vorbefund zweimal zu verwenden " & vbCrLf & _
"SELECT RANK() OVER (PARTITION BY pid ORDER BY zp,MPNr) rang, i.* FROM ( " & vbCrLf & _
"  -- 1) Anamnese:" & vbCrLf & _
"  SELECT a.pat_id pid, -1 MPNr, aufndat Zp, aufndat Bis, 'CSII' Thart, 'Anamnese: Insulinpumpe' Gru,  0 ia, x.absPos, x.StByte, 0 FeldNr FROM namen x LEFT JOIN anamnesebogen a USING (pat_id) " & vbCrLf & _
"  WHERE insulinpumpe<>0 AND (inpids IN('','0') OR x.pat_id IN (inpid)) " & vbCrLf & _
" UNION -- 2) Rezepte für Pumpenzubehör, könnten für ein Jahr Pumpentherapie versprechen: " & vbCrLf & _
"  SELECT Pat_id pid,-2 MPNr, zeitpunkt Zp,ADDDATE(zeitpunkt,365) bis,'CSII' Thart, feldinh Gru,foid ia,FID absPos,Form_id StByte, feldnr FROM formular x " & vbCrLf & _
"  WHERE form_abk IN ('rp','lar','prp','plar') AND feld IN ('medikament','txtMedKey','VerordnungsZeile') AND feldinh RLIKE 'reservoir|rapid d link|rap d li|rapid-d li|tenderl|sure t|paradigm|veo|animas|cartridge|t-slim|t:slim|variosoft|trusteel|autosoft|ypsopump|insigh|omnipod' AND NOT feldinh LIKE '%menveo%'" & vbCrLf & _
"  AND (inpids IN('','0') OR x.pat_id IN (inpid)) " & vbCrLf & _
" UNION -- 3) Insulinpläne " & vbCrLf & _
" SELECT pat_id pid,-3 MPNr,qdm zp,qdm bis,'ICT' Thart, MID(NAME,p) Gru,-2 ia,x.absPos,x.StByte,0 FROM (SELECT IF(p1>p2,p1,p2) p, b.* FROM (SELECT INSTR(b.name,'insulin') p1, INSTR(b.name,'spritz') p2, b.* FROM briefe b) b) x WHERE name RLIKE '(insulin|spritz).*(plan|schema|tabelle)' " & vbCrLf & _
"  AND (inpids IN('','0') OR x.pat_id IN (inpid)) " & vbCrLf & _
" UNION -- 4) Medikamentenplan " & vbCrLf & _
"  SELECT Pid,MPNr,Zp,Zp bis,Thart,gru,ia,abspos,stbyte,feldnr FROM ( " & vbCrLf & _
"   SELECT Pid,MPNr,Zp,Zp bis " & vbCrLf
 psql(3) = _
 "INSERT INTO therarten(pat_id,zp,mpnr,therart,insart,grund,abspos,aktzeit,stbyte)" & vbCrLf & _
 "SELECT pid,zp,mpnr,thart,ia,gru,abspos,aktzeit,stbyte FROM (" & vbCrLf & _
 "SELECT pid,zp,mpnr,thart, COALESCE(LAG(thart,1) OVER (PARTITION BY pid ORDER BY zp,MPNr),'') lthart" & vbCrLf & _
 ",ia,gru,abspos,NOW() aktzeit,stbyte FROM (" & vbCrLf & _
 "WITH dsort AS ( -- da zum Vergleich mit dem Vorbefund zweimal zu verwenden" & vbCrLf & _
 "SELECT RANK() OVER (PARTITION BY pid ORDER BY zp,MPNr) rang, i.* FROM (" & vbCrLf & _
 "  -- 1) Anamnese:" & vbCrLf & _
 "  SELECT a.pat_id pid, -1 MPNr, aufndat Zp, aufndat Bis, 'CSII' Thart, 'Anamnese: Insulinpumpe' Gru,  0 ia, x.absPos, x.StByte, 0 FeldNr FROM namen x LEFT JOIN anamnesebogen a USING (pat_id)" & vbCrLf & _
 "  WHERE insulinpumpe<>0 AND (inpid='0' OR x.pat_id IN (inpid))" & vbCrLf & _
 " UNION -- 2) Rezepte für Pumpenzubehör, könnten für ein Jahr Pumpentherapie versprechen:" & vbCrLf & _
 "  SELECT Pat_id pid,-2 MPNr, zeitpunkt Zp,ADDDATE(zeitpunkt,365) bis,'CSII' Thart, feldinh Gru,5 ia,FID absPos,Form_id StByte, feldnr FROM formular x" & vbCrLf & _
 "  WHERE form_abk IN ('rp','lar','prp','plar') AND feld IN ('medikament','txtMedKey','VerordnungsZeile') AND feldinh RLIKE 'reservoir|rapid d link|rap d li|rapid-d li|tenderl|flexl|sure t|paradigm|veo|animas|cartridge|t-slim|t:slim|variosoft|trusteel|autosoft|ypsopump|insigh|omnipod' AND NOT feldinh LIKE '%menveo%'" & vbCrLf & _
 "  AND (inpid='0' OR x.pat_id IN (inpid))" & vbCrLf & _
 " UNION -- 3) Rezepte für Nadeln oder Verzögerungsinsulin, die ICT versprechen:" & vbCrLf & _
 "  SELECT Pat_id pid,-2 MPNr, zeitpunkt Zp,zeitpunkt bis,'ICT' Thart, feldinh Gru,4 ia,FID absPos,Form_id StByte, feldnr FROM formular x" & vbCrLf & _
 "  WHERE form_abk IN ('rp','lar','prp','plar') AND feld IN ('medikament','txtMedKey','VerordnungsZeile') AND feldinh RLIKE 'fine|Nad|micro fi|[0-9] {0,1}mm|lantus|tresiba|levemir|basal|protaphan|semglee|abasaglar|semilente' AND NOT feldinh RLIKE 'fine {0,1}touch|Katheter|Paradigm|Mio|Flexl|Tender|d li|link|autosoft|Minimed|Sohlen|Oberarm|Fußbett|Schuh|Wanderh|Lancets fine|easy(-release| set)|quick {0,1}set|mmHg|insight|mylife|inset|insulinset|infusion|polster|szinti|dana|fexo|enadura|infektionsnadeln|port|magnes|medtronic|sicherheitslan|omnipod|orbisoft|orbit soft|nadellanz|schlauch|sure|tamponade|trusteel|varisoft|knoten|fine point|flex link|microlet fine|verkürz|vasofix|truesteel|sterile lanzetten|stahlnad|[68]0 {0,1}cm|alkohol|thin lanc|haut|TESTSTR|variosoft|Nadellänge|^BD Micro Fine Lancetten G 33 200 Stück$'" & vbCrLf & _
 "  AND (inpid='0' OR x.pat_id IN (inpid))" & vbCrLf & _
 " UNION -- 4) Insulinpläne" & vbCrLf & _
 " (SELECT pat_id pid,-3 MPNr,qdm zp,qdm bis,'ICT' Thart, MID(NAME,p) Gru,-2 ia,x.absPos,x.StByte,0 FROM (SELECT IF(p1>p2,p1,p2) p, b.* FROM (SELECT INSTR(b.name,'insulin') p1, INSTR(b.name,'spritz') p2, b.* FROM briefe b) b) x WHERE name RLIKE '(insulin|spritz).*(plan|schema|tabelle)'" & vbCrLf & _
 "  AND (inpid='0' OR x.pat_id IN (inpid))" & vbCrLf & _
 " UNION --  Insulinpläne als Formulare, angenommen wird mind. 14 Tage Gültigkeit für Therapieart" & vbCrLf & _
 " SELECT DISTINCT pat_id pid,-3 MPNr,zeitpunkt zp,zeitpunkt + INTERVAL 14 DAY bis,'ICT' Thart, CONCAT('Insulinplan ',(SELECT MAX(feldinh) FROM formular WHERE foid=x.foid AND feld='FAktuellesDatum')) Gru,-2 ia,1 absPos,1 StByte,0" & vbCrLf & _
 " FROM formular x WHERE feld='Eingabe1' AND (Nr = 3 AND Feldnr = 2) AND (inpid='0' OR x.pat_id IN (inpid))" & vbCrLf & _
 " GROUP BY pid, zp) " & vbCrLf
 psql(3) = psql(3) & _
 " UNION -- 5) Medikamentenplan" & vbCrLf & _
 "  SELECT Pid,MPNr,Zp,Zp bis,Thart,gru,ia,abspos,stbyte,feldnr FROM (" & vbCrLf & _
 "   SELECT Pid,MPNr,Zp,Zp bis" & vbCrLf
 psql(3) = psql(3) & _
 "       ,CASE WHEN SUM(pu) THEN 'CSII'" & vbCrLf & _
 "       WHEN SUM(obmzi) OR sum(iz)>=3 THEN -- wenn Mahlzeiteninsulin oder mind. 3 x Insulin/d" & vbCrLf & _
 "       CASE" & vbCrLf & _
 "        WHEN SUM(glp) THEN 'GLP1ICT'" & vbCrLf & _
 "        ELSE 'ICT'" & vbCrLf & _
 "       END" & vbCrLf & _
 "      WHEN SUM(iz)=2 THEN" & vbCrLf & _
 "       CASE" & vbCrLf & _
 "        WHEN SUM(glp) THEN 'GLP1Ins'" & vbCrLf & _
 "        ELSE 'CT'" & vbCrLf & _
 "       END" & vbCrLf & _
 "      WHEN SUM(iz)=1 THEN -- Insulinzahl pro Tag" & vbCrLf & _
 "       CASE" & vbCrLf & _
 "        WHEN SUM(glp) THEN 'GLP1Ins'" & vbCrLf & _
 "        WHEN SUM(oboad) THEN 'Komb'" & vbCrLf & _
 "        ELSE 'CT'" & vbCrLf & _
 "       END" & vbCrLf & _
 "      ELSE -- WHEN SUM(iz)=0 THEN" & vbCrLf & _
 "       CASE" & vbCrLf & _
 "        WHEN SUM(glp) THEN 'GLP1'" & vbCrLf & _
 "        WHEN SUM(oboad) THEN 'OAD'" & vbCrLf & _
 "        ELSE 'Diät'" & vbCrLf & _
 "       END" & vbCrLf
 psql(3) = psql(3) & _
 "     END Thart" & vbCrLf & _
 "   ,GROUP_CONCAT(DISTINCT gru SEPARATOR '') gru,MAX(ia) ia,abspos,stbyte,MIN(feldnr) feldnr -- Grund, Insulinart" & vbCrLf & _
 "   FROM (" & vbCrLf & _
 "    SELECT Pid,MPNr,Zp,Med,Pu,oboad,IF(obin,ezm,0) iz,ezm AND ia=1 obmzi" & vbCrLf & _
 "     ,IF(eztm,glp,0) glp,IF(ezm,obin,0) ins,IF(ezm,ia,0) ia" & vbCrLf & _
 "     ,IF(pu||if(eztm,glp,0)||oboad||(ezm&&ia=1)||if(obin,ezm,0),CONCAT('/',Med,' '),'') gru" & vbCrLf & _
 "     ,IF(pu||if(eztm,glp,0)||oboad||(ezm&&ia=1)||if(obin,ezm,0),Feldnr,NULL) FeldNr,absPos,StByte" & vbCrLf & _
 "    FROM (" & vbCrLf & _
 "     SELECT pid,MPNr,zp,Med,pu,ez,wglp,ohneE,IF(ez,oad,0)oboad,glp,obin,ia,FeldNr,absPos,StByte" & vbCrLf & _
 "      ,IF(ez,ez,ohnee) ezm -- Eintragszahl modifziert" & vbCrLf & _
 "      ,IF(ez,ez,wglp&&ohneE) eztm -- Eintragszahl teilmodifiziert" & vbCrLf & _
 "     FROM (" & vbCrLf & _
 "      SELECT x.Pat_id pid,x.MPNr MPNr,x.Zeitpunkt Zp,x.Medikament Med,ma.puzu<>0 pu" & vbCrLf & _
 "       ,MAX((COALESCE(x.mo,'')<>'')+(COALESCE(x.mi,'')<>'')+(COALESCE(x.nm,'')<>'')+(COALESCE(x.ab,'')<>'')+(COALESCE(x.zn,'')<>'')+if(glp1<>0 AND x.Medanfang RLIKE 'OZEMPIC|TRULICITY|bydureon|victoza|mounjaro',1,0)) ez" & vbCrLf & _
 "       ,glp1<>0 AND x.MedAnfang RLIKE 'OZEMPIC|TRULICITY|bydureon|mounjaro' wglp -- Wochen-GLP-1" & vbCrLf & _
 "       ,pzn<>0 AND concat(x.Bemerkung,' ',x.Grund) NOT RLIKE 'Pause|abgesetzt|beendet|zur Zeit nicht' ohneE -- ohne Eintrag im neuen Medplan, aber vermutlich angewandt" & vbCrLf & _
 "       ,ma.glib<>0 OR ma.metf<>0 OR ma.gluci<>0 OR ma.shglin<>0 OR ma.dpp4<>0 OR ma.sglt2<>0 OR ma.sonstad<>0 oad" & vbCrLf & _
 "       ,glp1<>0 glp" & vbCrLf & _
 "       ,ma.ins<>0 OR ma.anal<>0 obin -- ob Insulin" & vbCrLf & _
 "       ,IF(insart='' OR ISNULL(insart),0,insart) ia -- Insulinart: 1= schnell, 2 = langsam, 3 = Misch" & vbCrLf & _
 "       ,x.FeldNr,x.absPos,x.StByte -- zur Zeit eher überflüssige Felder" & vbCrLf & _
 "      FROM wmedplan x LEFT JOIN medarten ma ON x.medanfang= ma.medikament" & vbCrLf & _
 "      WHERE (inpid='0' OR x.pat_id IN (inpid))" & vbCrLf
 psql(3) = psql(3) & _
 "      GROUP BY x.pat_id,mpnr,x.zeitpunkt,ma.id -- z.B. versch.Toujeo-Zeilen für v. Zuckerwerte" & vbCrLf & _
 "     ) i" & vbCrLf & _
 "    ) i" & vbCrLf & _
 "   ) i GROUP BY pid,MPNr,zp -- 13.12.21: gleiche MPNr für versch. Zeitpunkte! wohl durch HB_-Import" & vbCrLf & _
 "  ) i" & vbCrLf & _
 " ) i" & vbCrLf & _
 ") -- with ...; letzte Therapieart, letzter Zeitpunkt: " & vbCrLf & _
 "SELECT d.*" & vbCrLf & _
 ", COALESCE(ldso.thart,'') lathart -- letzte andere Therapieart" & vbCrLf & _
 ", RANK() over (PARTITION BY pid ORDER BY zp,mpnr) thrang" & vbCrLf & _
 ", COALESCE(LAG(d.zp,1) OVER (PARTITION BY pid ORDER BY zp,MPNr),'1900-01-01') lzp" & vbCrLf & _
 ", COALESCE(LAG(d.ia,1) OVER (PARTITION BY pid ORDER BY zp,MPNr),'-10') lia" & vbCrLf & _
 "FROM dsort d LEFT JOIN dsort ldso -- letztes (anderes) dsort" & vbCrLf & _
 " ON ldso.pid=d.pid AND ldso.rang=(SELECT MAX(rang) FROM dsort WHERE pid=d.pid AND rang<d.rang AND thart<>d.thart)" & vbCrLf & _
 " WHERE NOT EXISTS (SELECT 1 FROM dsort WHERE pid=d.pid AND rang<d.rang AND (zp<d.zp AND bis>=d.zp)) -- Gültigkeitsende wirken lassen (bes. Pumpenrezept)" & vbCrLf & _
 ") i -- nur Wechsel anzeigen, nicht von CSII/ICT auf Diät, nicht im Karenzzeitraum (1a nach Pumpenrezept oder 92 Tage nach Insulinplan):" & vbCrLf & _
 "WHERE false OR thart<>lathart AND NOT (" & vbCrLf & _
 "  thart='Diät' AND lathart IN ('CSII','ICT','GLP1ICT') AND NOT EXISTS (SELECT 1 FROM sws WHERE pat_id=i.pid AND voret BETWEEN i.lzp AND i.zp))" & vbCrLf & _
 "  AND NOT (lia=-2 AND ia=2 AND NOT (thart<>lathart AND thart IN ('GLP1','GLP1Ins','GLP1ICT')) AND zp BETWEEN lzp AND ADDDATE(lzp,92))" & vbCrLf & _
 "  AND NOT (ia=4 and lathart<>'CSII')" & vbCrLf & _
 "  AND NOT (lathart='CSII' AND ia<>4)" & vbCrLf & _
 "GROUP BY pid,zp,thart" & vbCrLf & _
 "ORDER BY pid,zp,MPNr" & vbCrLf & _
 ") i WHERE thart<>lthart" & vbCrLf
 psql(3) = psql(3) & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf & _
 "" & vbCrLf

 psql(4) = "SELECT @vzahl vzahl, ROW_COUNT() zahl; "
' psql(5) = "UPDATE anamnesebogen x SET ther1=(SELECT therart FROM therarten WHERE pat_id=x.pat_id ORDER BY zp, mpnr LIMIT 1), therakt =(SELECT therart FROM therarten WHERE pat_id=x.pat_id ORDER BY zp DESC, mpnr DESC LIMIT 1) WHERE inpids IN('','0') OR x.pat_id IN (inpid);"
 psql(5) = "UPDATE anamnesebogen x SET ther1=(SELECT therart FROM therarten WHERE pat_id=x.pat_id ORDER BY zp, mpnr LIMIT 1), therakt =(SELECT therart FROM therarten WHERE pat_id=x.pat_id ORDER BY zp DESC, mpnr DESC LIMIT 1) WHERE (inpid='0' OR x.pat_id IN (inpid));"
End Function ' therinit

' 10/22: die folgende Funktion muss so umständlich eingerichtet werden, da der Aufruf von "call fuellThaP"
' unter MariaDB 10.9 mit ca. 80% Wahrscheinlichkeit den Server crasht (ähnliches im Netz)
' gemeinsame Algorithmusdefinition in therinit
' in alleSpeichern, doViewsErstellen, testTab, Therapieartenwechsel_click, rufThFestleg, theraktakt
Public Function TheraErmitt&(pids$, Optional ByRef vzahl&, Optional Position$)
 Dim iru&, rs As ADODB.Recordset, rAf& ' , inpids$
 syscmd 4, "Ermittle Therapiearten für Pat. " & pids & Position
' Lese.ProgStart
 ComTrans
 Call therinit
' Select Case Left$(pids, 1)
'  Case """", "'": pids = Mid$(pids, 2)
' End Select
' Select Case Right$(pids, 1)
'  Case """", "'": pids = Left$(pids, Len(pids) - 1)
' End Select
  pids = REPLACE$(REPLACE$(pids, "'", ""), """", "")
' inpids = "'" & pids & "'"
 For iru = 1 To 5 ' ohne group_concat_max_len
'  sql = replace$(psql(iru), "inpid", pids) ' FIND_IN_SET
  Select Case iru
   Case 1, 3, 5
    Select Case pids
     Case "", "''", "0", "'0'"
      sql = REPLACE$(REPLACE$(psql(iru), "inpid=", "'0'="), "(inpid)", "(SELECT pat_id FROM namen)")
     Case Else
      sql = REPLACE$(REPLACE$(psql(iru), "inpid=", "'inpid'="), "inpid", pids)
    End Select
'   Case 1
'    sql = REPLACE(replace$(psql(iru), "FIND_IN_SET(pat_id,inpid)>0", "pat_id IN (" & pids & ")"), "inpid", pids)
'   Case 3, 5
'    Select Case pids
'     Case "", "0":
'      sql = replace$(psql(iru), "inpid IN('','0') OR FIND_IN_SET(x.pat_id,inpid)>0", "true")
'     Case Else
'      sql = replace$(psql(iru), "inpid IN('','0') OR FIND_IN_SET(x.pat_id,inpid)>0", "x.pat_id IN (" & pids & ")")
   Case Else
    sql = psql(iru)
  End Select
 syscmd 4, "Ermittle Therapiearten für Pat. " & pids & " (Runde: " & iru & ")" & Position
  myFrag rs, sql, adOpenUnspecified, Nothing, adLockReadOnly, 15000, rAf
  If iru = 1 Then
   vzahl = rAf ' Zahl der vorherigen Datensätze in in Therarten
  ElseIf iru = 4 Then
   TheraErmitt = rs!Zahl ' Zahl der aktuellen Datensätze in in Therarten
  End If ' iru = 1
 Next iru
 ComTrans
 syscmd 5
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in TheraErmitt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' TheraErmitt


'Public lies.obmysql%
#If False Then
Function ICDZeit() ' kommt nirgends vor
 Dim rsAnam As New ADODB.Recordset
 Const LDeP$ = FName
 On Error GoTo fehler
' Call dtbInit
 DoCmd.Close acForm, LDeP, acSaveYes
 DoCmd.Close acQuery, LDeP, acSaveYes
' DoCmd.DeleteObject acQuery, LDeP
' SELECT `Pat_id`, `Nachname`, `Vorname`, `Diabetestyp`, `Diabetes seit`, "       " AS ICD, true AS ob, "                 " AS Begr1, "                 " AS Begr2, "                 " AS Begr3, "                 " AS Begr4 FROM _Anamnesebogen;
 DtbQuerydefsDelete LDeP
' SET rsAnam = TabÖff("Anamnesebogen")
' rsAnam.Open "SELECT * FROM `anamnesebogen`", DBCn, adOpenDynamic, adLockReadOnly
 myFrag rsAnam, "SELECT * FROM `anamnesebogen`"
 If Not rsAnam.BOF Then
  Do While Not rsAnam.EOF
'   rsAnam.Edit
   If rsAnam!Größe > 3 Then rsAnam!Größe = rsAnam!Größe / 100
   rsAnam!ob = 0
   If rsAnam!Größe <> 0 Then
    If rsAnam!Gewicht * IIf(rsAnam!Gewicht < 3, 100, 1) / rsAnam!Größe / rsAnam!Größe >= 29 And rsAnam!Gewicht / rsAnam!Größe / rsAnam!Größe <> 0 And InStrB(rsAnam!Diagnosen, "E66.99") = 0 Then
     rsAnam!ob = -1
    End If
   End If
   rsAnam.Update
   rsAnam.Move 1
  Loop
 End If
 '  ORDER BY gewicht/größe/größe*10000 desc
 Call DtbCreateQueryDef(LDeP, "SELECT * FROM (SELECT Pat_id, Nachname + "", "" + Vorname AS Name,Diabetestyp,`Diabetes seit`, ""E66.99"" AS ICD,ob,round(gewicht * iIF(gewicht<3,100,1)/größe/größe,2) AS Begr1,"""" AS begr2, """" AS begr3, """" AS begr4 FROM `anamnesebogen`) ORDER BY ob")
 DoCmd.OpenForm LDeP
 Forms(LDeP).Caption = "ICD-Auswahl Übergewicht"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ICDZeit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ICDZeit
#End If

Private Function cmd$(sql$, obAcc%)
 If obAcc Then
  cmd = sql
'  Cmd = replace$(Cmd, "`", vns)
  cmd = REPLACE$(cmd, "date_add", vNS)
  cmd = REPLACE$(cmd, "DATE_SUB", vNS)
  cmd = REPLACE$(cmd, ",interval", "-")
  cmd = REPLACE$(cmd, ", interval", "+")
  cmd = REPLACE$(cmd, "INTERVAL ", " ")
  cmd = REPLACE$(cmd, "day)", ")")
  cmd = REPLACE$(cmd, "month)", "*30)")
  cmd = REPLACE$(cmd, "CAST(", "cdate(int(")
  cmd = REPLACE$(cmd, "STR_TO_DATE(", "cdate(")
  cmd = REPLACE$(cmd, ",'%d.%m.%Y'", vNS)
  cmd = REPLACE$(cmd, "as date", ")")
  cmd = REPLACE$(cmd, "_utf8", " ")
  cmd = REPLACE$(cmd, "concat", vNS)
  cmd = REPLACE$(cmd, "substr(", "MID(")
  cmd = REPLACE$(cmd, "¡", " & ")
  cmd = REPLACE$(cmd, "intaccdatemy", "int")
  cmd = REPLACE$(cmd, "intacc", "int")
  cmd = REPLACE$(cmd, "SELECTmy", "((")
  cmd = REPLACE$(cmd, "divmy", "/")
  cmd = REPLACE$(cmd, "to_days", "int")
'  Cmd = replace$(replace$(Cmd, "CAST(", vns), " As Date)", vns)
  cmd = REPLACE$(cmd, "SUBDATE(", "(")
  cmd = REPLACE$(cmd, "adddate(", "(")
 '  Call cat.Views.Append(vn(i), Cmd)
  cmd = REPLACE$(cmd, "STR_TO_DATE(", vNS)
  cmd = REPLACE$(cmd, ",'#%m/%d/%Y#')", ")")
  cmd = REPLACE$(cmd, ",""#%m/%d/%Y#"")", ")")
'  Call Shell(App.Path & "\viewserst.exe """ & Forms(0).dbv.Datei & """ """ & QName & """ """ & Cmd & """")
 Else
  cmd = sql
  cmd = REPLACE$(cmd, "SELECTmy", "SELECT")
  cmd = REPLACE$(cmd, "divmy", "div")
  cmd = REPLACE$(cmd, "¡", ",")
  cmd = REPLACE$(cmd, "intaccdatemy", "date")
  cmd = REPLACE$(cmd, "intacc", vNS)
  cmd = REPLACE$(cmd, "cdate(", "(")
  cmd = REPLACE$(cmd, "first(", "(")
'  Call myEFrag("DROP VIEW " & ifexists & " `" & QName & "`;")
'  Call myEFrag("CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`" & Forms(0).dbv.uid & "`@`%` SQL SECURITY DEFINER VIEW `" & QName & "` AS " & Cmd)
 End If
End Function ' Cmd

Function SqlU(sql$, obmy%) As CString
 Set SqlU = New CString
 SqlU = sql
 If obmy Then
  SqlU.REPLACE "¡", ","
  SqlU.REPLACE "intaccdatemy", "DATE"
  Call EliminierWortplusZahl(SqlU, "TOP ")
  
  SqlU.REPLACE "SELECTmy", "SELECT"
  SqlU.REPLACE "divmy", "DIV"
  SqlU.REPLACE "intacc", vNS
  SqlU.REPLACE "cdate(", "("
  SqlU.REPLACE "first(", "("
'  Call myEFrag("DROP VIEW " & ifexists & " `" & QName & "`;")
'  Call myEFrag("CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`" & Forms(0).dbv.uid & "`@`%` SQL SECURITY DEFINER VIEW `" & QName & "` AS " & Cmd)
 Else
  SqlU.REPLACE "¡", " & "
  SqlU.REPLACE "concat", vNS
  SqlU.REPLACE "intaccdatemy", "int"
  Call EliminierWortplusZahl(SqlU, "limit ")
 
  SqlU.REPLACE "date_add", vNS
  SqlU.REPLACE "DATE_SUB", vNS
  SqlU.REPLACE ",interval", "-"
  SqlU.REPLACE ", interval", "+"
  SqlU.REPLACE "INTERVAL ", " "
  SqlU.REPLACE "day)", ")"
  SqlU.REPLACE "month)", "*30)"
  SqlU.REPLACE "CAST(", "cdate(int("
  SqlU.REPLACE "STR_TO_DATE(", "cdate("
  SqlU.REPLACE ",'%d.%m.%Y'", vNS
  SqlU.REPLACE "as date", ")"
  SqlU.REPLACE "_utf8", " "
  SqlU.REPLACE "substr(", "MID("
  SqlU.REPLACE "¡", " & "
  SqlU.REPLACE "intacc", "int"
  SqlU.REPLACE "SELECTmy", "(("
  SqlU.REPLACE "divmy", "/"
  SqlU.REPLACE "to_days", "int"
'  Cmd = replace$(replace$(Cmd, "CAST(", vns), " As Date)", vns)
  SqlU.REPLACE "SUBDATE(", "("
  SqlU.REPLACE "adddate(", "("
 '  Call cat.Views.Append(vn(i), Cmd)
  SqlU.REPLACE "STR_TO_DATE(", vNS
  SqlU.REPLACE ",'#%m/%d/%Y#')", ")"
  SqlU.REPLACE ",""#%m/%d/%Y#"")", ")"
 End If
End Function ' SqlUmw

Sub EliminierWortplusZahl(ByRef SqlU As CString, ByRef Wort$)
 Dim pos&, p2&, sp$, p2anf&
 pos = SqlU.Instr(Wort)
 If pos <> 0 Then
  sp = SqlU.Left(pos - 1)
  p2 = pos + Len(Wort)
  p2anf = p2
  Do
   If Not IsNumeric(SqlU.Mid(p2, 1)) Then Exit Do
   p2 = p2 + 1
  Loop
  If p2anf <> p2 Then SqlU = SqlU.Left(pos - 1) & SqlU.Mid(p2)
 End If
End Sub ' EliminierWortplusZahl(ByRef SqlU As CString, ByRef Wort$)

Function DtbCreateQueryDef$(QName$, sql$)
 Dim csql As CString
 If sql <> vNS Then
 Set csql = SqlU(sql, ((LVobMySQL)))
 If Not LVobMySQL Then
  Dim cat As New ADOX.Catalog
'  Dim Cmd$ ' As New ADODB.Command
  Dim V As ADOX.view
'  SET Cmd.ActiveConnection = DBCn
  Set cat.ActiveConnection = DBCn
  On Error GoTo fehler
'  SET Cmd = Nothing
'  Cmd.CommandText =
 '  Call cat.Views.Append(vn(i), Cmd)
'  Call Shell(App.Path & "\viewserst.exe """ & Forms(0).dbv.Datei & """ """ & QName & """ """ & Cmd & """")
' hier ist wegen eines Fehlers in ADOX wirklich noch DAO nötig
'#If False THEN
  Dim QTb ' AS dao.Database
'  SET QTb = DAO.OpenDatabase(Forms(0).dbv.Datei)
  Set QTb = CreateObject("DAO.DBEngine.36").OpenDatabase(Forms(0).dbv.Datei)
  On Error Resume Next
'  Call cat.Views.Delete(QName)
  QTb.QueryDefs.Delete QName
  On Error GoTo fehler
  Call QTb.CreateQueryDef(QName, csql) '.CommandText)
  Set QTb = Nothing
'#END IF
 Else
  On Error GoTo fehler
  If LVobMySQL And LenB(ifexists) = 0 Then Zinit (LVobMySQL)
  Call myEFrag("DROP TABLE " & ifexists & " `" & QName & "`;")
  Call myEFrag("DROP VIEW " & ifexists & " `" & QName & "`;")
  On Error GoTo fehler
  Dim cvrs As Recordset
  
  Call myEFrag("CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`" & Forms(0).dbv.uid & "`@`%` SQL SECURITY DEFINER VIEW `" & QName & "` AS " & csql)
  Set cvrs = myEFrag("SHOW TABLES WHERE `tables_in_" & DefDB(DBCn) & "` LIKE '" & QName & "'")
  If cvrs.BOF Then
   Lese.Ausgeb QName & " konnte nicht erstellt werden.", True
  Else
'   Debug.Print QName & " gibts."
  End If
 End If
 DtbCreateQueryDef = csql
 End If ' sql = vns
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If InStrB(Err.Description, "gone away") <> 0 Then
 Call DBCnOpen
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DtbCreateQueryDef/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DtbCreateQueryDef

Function do_Labor_Click(frm As Form)
 On Error GoTo fehler
 Call do_Click_Vorbereit
 'sql = "SELECT dl.`Pat_ID` AS Pat_id, `Namen`.`Nachname` AS Nachname, `Namen`.`Vorname` AS Vorname, dl.`ZeitPunkt`, dl.`DokName`, dl.`AbsPos`, dl.`AktZeit`, dl.`DokPfad`, dl.`dokgroe`, `br_abgehakt`.`abgehakt` " & _
 "FROM `namen` INNER JOIN ((SELECT * FROM `dokumente` WHERE dokname LIKE ""*labor*"") AS dl LEFT JOIN `br_abgehakt` ON dl.`DokPfad`=`br_abgehakt`.`DokPfad`) ON `Namen`.`Pat_ID`=`Dl`.`Pat_ID` " & _
 "ORDER BY `Namen`.`Nachname`, `Namen`.`Vorname`, dl.`ZeitPunkt`;"
 sql = "SELECT dl.Pat_ID AS Pat_id, Nachname, Vorname, ZeitPunkt, dl.Name DokName, dl.AbsPos, dl.AktZeit, dl.Pfad DokPfad, dokgroe, abgehakt " & _
 "FROM (briefe dl LEFT JOIN `br_abgehakt` da ON dl.Pfad=da.DokPfad) LEFT JOIN `namen` na ON na.Pat_ID=dl.Pat_ID " & _
 "WHERE Name LIKE '%labor%' AND dl.pat_id = " + CStr(frm!vTextB(1)) + _
 " ORDER BY Nachname, Vorname, ZeitPunkt DESC;"

 Call DtbCreateQueryDef("LaborDokumente eP", sql)
 DoCmd.OpenForm "LaborDokumente eP"
 Forms![LaborDokumente ep].Caption = "LaborDokumente zu " & IIf(IsNull(frm!Nachname), frm!Pat_ID, frm!Nachname) & " " & IIf(IsNull(frm.Vorname), vNS, frm.Vorname)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Labor_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Labor_Click(frm AS Form)

#If mitab Then
Function do_DokDown(frm As Form) ' kommt nirgends vor
 Dim Oneu$, rDok As New ADODB.Recordset
 On Error GoTo fehler
' Call dtbInit
 If PcDokPfad = vNS Then Call getDokPfad
 If FSO Is Nothing Then Set FSO = CreateObject("Scripting.FileSystemObject")
 Oneu = frm.Nachname + "_" + frm.Vorname + "_" + Format$(frm.GebDat, "dd/mm/yyyy")
 Call VerzPrüf(pVerz & Oneu)
 sql = "SELECT """ + frm.Nachname + """ Nachname,""" + frm.Vorname + """ Vorname, b.* FROM briefe b WHERE name NOT LIKE ""%labor%"" AND name NOT LIKE ""%heckliste%"" AND name NOT LIKE ""%namnese%"" AND pat_id = " + CStr(frm!Pat_ID) + " ORDER BY zeitpunkt DESC"
' SET rDok = Dtb.OpenRecordset(sql, dbOpenDynaset)
 myFrag rDok, sql
 On Error Resume Next
 Do While Not rDok.EOF
  FSO.CopyFile vNS + REPLACE$(REPLACE$(LCase$(rDok!DokPfad), "$\turbomed\dokumente", PcDokPfad), "^", vNS) & vNS, vNS & pVerz '+ Oneu + "\" + rDok!DokName + vns
  rDok.Move 1
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_DokDown/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_DokDown
#End If

' in AnBog
Function do_Dokumente_Click(frm As Form)
 On Error GoTo fehler
 Call do_Click_Vorbereit
 sql = "SELECT """ + frm.Nachname + """ nachname,""" + frm.Vorname + """ vorname, b.* FROM " & "`" & QMdbAkt & "`" & ".briefe b WHERE name NOT LIKE ""%labor%"" AND name NOT LIKE ""%heckliste%"" AND name NOT LIKE ""%namnese%"" AND NOT (name LIKE ""%BZ%"" OR name LIKE ""%RR%"" OR name LIKE ""%blutdr%"") AND pat_id = " & CStr(frm!Pat_ID) + " ORDER BY zeitpunkt DESC"
 Call DtbCreateQueryDef("LaborDokumente eP", sql)
' DoCmd.OpenForm "LaborDokumente eP"
 DoCmd.OpenForm "LaborDokumente"
' Forms![LaborDokumente ep].Caption = "Sonstige Dokumente zu " + frm.Nachname & " " & frm.Vorname
 Forms![LaborDokumente].Caption = "Sonstige Dokumente zu " + frm.Nachname & " " & frm.Vorname
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Dokumente_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Dokumente_Click(frm AS Form)

' in vCommandB_Click
Function do_Briefe_Click(frm As Form)
 On Error GoTo fehler
 Call do_Click_Vorbereit
 sql = "SELECT """ + frm.Nachname + """ nachname,""" + frm.Vorname + """ vorname, pfad DokPfad, name DokName, art DokArt, b.* FROM `" + QMdbAkt + "`.`briefe` b WHERE pat_id = " + CStr(frm!Pat_ID) + " ORDER BY zeitpunkt DESC"
 Call DtbCreateQueryDef("LaborDokumente eP", sql)
 DoCmd.OpenForm "LaborDokumente eP", acNormal, , "17=17"
 Forms![LaborDokumente ep].Caption = "Briefe zu " + frm.Nachname & " " & frm.Vorname
 ' irfan
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Briefe_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Briefe_Click(frm AS Form)

' in vCommandb_Click
Function do_BZKurven_Click(frm As Form)
 On Error GoTo fehler
 Call do_Click_Vorbereit
 Call DtbCreateQueryDef("LaborDokumente eP", "SELECT """ + frm.Nachname + """ nachname,""" + frm.Vorname + """ vorname, b.* FROM `" + QMdbAkt + "`.briefe b WHERE (ame LIKE ""%BZ%"" OR name LIKE ""%RR%"" OR name LIKE ""%blutdr%"") AND pat_id = " + CStr(frm!Pat_ID) + " ORDER BY zeitpunkt DESC")
 DoCmd.OpenForm "LaborDokumente eP", acNormal, , "17=17"
 Forms![LaborDokumente ep].Caption = "Blutzuckerkurven zu " + frm.Nachname & " " & frm.Vorname
 ' irfan
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_BZKurven_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_BZKurven_Click(frm AS Form)

' in vCommandB_Click
Function do_AugenBefunde_Click(frm As Form)
 On Error GoTo fehler
 Call do_Click_Vorbereit
 Call DtbCreateQueryDef("LaborDokumente eP", "SELECT """ + frm.Nachname + """ nachname,""" + frm.Vorname + """ vorname, zeitpunkt, aktzeit, name dokname, pfad dokpfad, dokart, pat_id FROM `" + QMdbAkt + "`.briefe b WHERE (name LIKE ""%augenb`!l`%"" OR name LIKE ""%augen`aä`rzt%"" OR name LIKE ""%aa%"") AND pat_id = " + CStr(frm!Pat_ID) + _
                                                " UNION SELECT """ + frm.Nachname + """ nachname,""" + frm.Vorname + """ vorname, zeitpunkt, aktzeit, inhalt DokName, art dokpfad, art DokArt, pat_id FROM `" + QMdbAkt + "`.`eintraege` WHERE (inhalt LIKE ""%augenb`!l`%"" OR inhalt LIKE ""%augen`aä`rzt%"" OR inhalt LIKE ""% aa`!g`%"") AND pat_id = " + CStr(frm!Pat_ID) + " ORDER BY zeitpunkt DESC")
 DoCmd.OpenForm "LaborDokumente eP", acNormal, , "17=17"
 Forms![LaborDokumente ep].Caption = "Augendokumente zu " + frm.Nachname & " " & frm.Vorname
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_AugenBefunde_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_AugenBefunde_Click(frm As Form)

Function DtbQuerydefsDelete(view$)
 Dim cat As New ADOX.Catalog
 cat.ActiveConnection = DBCn
 On Error Resume Next
 Call cat.Views.Delete(view)
End Function ' DtbQuerydefsDelete

Function do_Click_Vorbereit()
 Const LDeP$ = "LaborDokumente eP"
 On Error GoTo fehler
#If obDAO Then
 Call dtbInit
#End If
 DoCmd.Close acForm, LDeP, acSaveYes
 DoCmd.Close acQuery, LDeP, acSaveYes
' DoCmd.DeleteObject acQuery, LDeP
 DtbQuerydefsDelete LDeP
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Click_Vorbereit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Click_Vorbereit()

Function do_HAC(HA_Ausw$, nr%)
 ReDim Preserve HANrBf(2)
 If HA_Ausw Like "##/#####*" Then
  If REPLACE$(HA_Ausw, "/", vNS) <> HANrBf(nr) Then
   HANrBf(nr) = REPLACE$(HA_Ausw, "/", vNS)
   Call Sound(WinDir + "\media\Windows XP-Batterie niedrig.wav")
  End If
 End If
End Function ' do_HAC(HA_Ausw$, nr%)

Function dodo_u_Click(frm As AnBog, nr)
 Dim rfDE As New ADODB.Recordset, i&
 Dim rAf&
 On Error GoTo fehler
 If UStumm Then Exit Function
' Debug.Print frm.Controls("U" + CStr(nr)).Value
' rfDE.Seek "=", frm!Pat_id, MDIICD(Nr), MDIDiag(Nr)
 Dim icdlike$
 If MDIICD(nr) <> vNS Then
  If MDIICD(nr) Like "E1*" Or MDIICD(nr) Like "L89*" Then
   icdlike = Left$(MDIICD(nr), InStr(MDIICD(nr), ".")) & ".*"
  Else
   icdlike = MDIICD(nr)
   For i = 1 To 2
    If Right$(icdlike, 1) = "V" Then icdlike = Left$(icdlike, Len(icdlike) - 1)
    If Right$(icdlike, 1) = "Z" Then icdlike = Left$(icdlike, Len(icdlike) - 1)
   Next i
   icdlike = icdlike & "[VZ]?"
  End If
  ' frm.anaRS!Pat_id
'  Call rfDE.Open("SELECT * FROM `fuerdiagexp` WHERE pat_id = " & Pat_id & " AND icd RLIKE '" & icdlike & "' AND diagnose LIKE '%" & LTrim$(replace$(LTrim$(replace$(MDIDiag(nr), "V.a.", vNS)), "Z.n.", vNS)) & "'", DBCn, adOpenDynamic, adLockOptimistic)
  myFrag rfDE, "SELECT * FROM `fuerdiagexp` WHERE pat_id = " & Pat_ID & " AND icd RLIKE '" & icdlike & "' AND diagnose LIKE '%" & LTrim$(REPLACE$(LTrim$(REPLACE$(MDIDiag(nr), "V.a.", vNS)), "Z.n.", vNS)) & "'"
  UStumm = True
  frm.Va = 0
  frm.Zn = 0
  frm.Xtra = vNS
  If frm.Controls("vOptionB")(nr).Value <> 0 Then
   If rfDE.BOF Then
    If MDIICD(nr) Like "*V*" Then frm.Va = 1 Else If MDIICD(nr) Like "*Z*" Then frm.Zn = 1
   Else
    If rfDE!ICD Like "*V*" Then frm.Va = 1 Else If rfDE!ICD Like "*Z*" Then frm.Zn = 1
   End If
   For i = frm.Xtra.ListCount - 1 To 0 Step -1
    frm.Xtra.RemoveItem i
   Next i
   If MDIICD(nr) Like "E1*" Then
    frm.Xtra.AddItem ".01 mit Koma"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 0
    frm.Xtra.AddItem ".11 mit Ketoazidose"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 1
    frm.Xtra.AddItem ".21 mit Nierenkomplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 2
    frm.Xtra.AddItem ".31 mit Augenkomplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 3
    frm.Xtra.AddItem ".41 mit neurologischen Komplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 4
    frm.Xtra.AddItem ".51 mit peripheren vaskulären Komplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 5
    frm.Xtra.AddItem ".61 mit sonstigen näher bezeichneten Komplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 6
    frm.Xtra.AddItem ".73 mit multiplen Komplikationen ohne Fußsyndrom"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 7
    frm.Xtra.AddItem ".75 mit multiplen Komplikationen mit Fußsyndrom"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 8
    frm.Xtra.AddItem ".81 mit nicht näher bezeichneten Komplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 9
    frm.Xtra.AddItem ".91 ohne Komplikationen"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 10
   ElseIf MDIICD(nr) Like "L89*" Then
    frm.Xtra.AddItem ".17 Wagner 0 Ferse"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 0
    frm.Xtra.AddItem ".18 Wagner 0 ohne Ferse"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 1
    frm.Xtra.AddItem ".27 Wagner 1 Ferse (oberfl., Infektion nur im Wundbereich)"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 2
    frm.Xtra.AddItem ".28 Wagner 1 ohne Ferse (oberfl., Infektion nur im Wundbereich)"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 3
    frm.Xtra.AddItem ".37 Wagner 2 Ferse (tief, gelenknah)"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 4
    frm.Xtra.AddItem ".38 Wagner 2 ohne Ferse (tief, gelenknah)"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 5
    frm.Xtra.AddItem ".47 Wagner 3 Ferse (Teilnekrose)"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 6
    frm.Xtra.AddItem ".48 Wagner 3 ohne Ferse (Teilnekrose)"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 7
    frm.Xtra.AddItem ".97 Wagner nicht näher bezeichnet Ferse"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 8
    frm.Xtra.AddItem ".98 Wagner nicht näher bezeichnet ohne Ferse"
    frm.Xtra.ItemData(frm.Xtra.NewIndex) = 9
   End If
  End If
  UStumm = False
  If rfDE.BOF And frm.Controls("vOptionB")(nr).Value <> 0 Then
'   MsgBox "Stop in dodo_u_Click:" & vbCrLf & "MDIDiag(nr)=""""" & vbCrLf & "Nr: " & Nr
   If MDIDiag(nr) = vNS Then
    MsgBox "Fehler in dodo_u_Click: MDIDiag(Nr) = vNS"
    Stop
   End If
   ' frm.anaRS!Pat_id
   InsKorr DBCn, "INSERT INTO `fuerdiagexp`(pat_id,name,icd,diagnose,nurquart,zeitpunkt) VALUES(" & Pat_ID & ",'" & frm.anaRS!Nachname & " " & frm.anaRS!Vorname & "','" & MDIICD(nr) & "','" & MDIDiag(nr) & "'," & IIf(MDIICD(nr) = "O24.4" Or (MDIICD(nr) Like "L89*" And Not MDIICD(nr) Like "L89.1*"), 1, 0) & "," & DatFor_k(Now()) & ")", rAf
  ElseIf frm.Controls("vOptionB")(nr).Value = False Then
   ' frm.anaRS!Pat_id
   myEFrag "DELETE FROM `fuerdiagexp` WHERE pat_id = " & Pat_ID & " AND icd = '" & MDIICD(nr) & "' AND diagnose = '" & MDIDiag(nr) & "'", rAf
  End If
 End If ' MDIICD(nr) <> vns THEN
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
If InStrB(Err.Description, "Transaction level 'READ-COMMITTED'") <> 0 Then
 myEFrag "SET SESSION TRANSACTION ISOLATION LEVEL REPEATABLE READ", rAf
 Resume
End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dodo_u_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dodo_u_Click

Function do_Diagnosen_Reset(Optional frm As Form)
 Dim rfDE As New ADODB.Recordset, rfDEA As New ADODB.Recordset, F As ADODB.Field, erg&
 Dim archiviert As Date
 erg = MsgBox("Wollen Sie wirklich alle angekreuzten Diagnosen zurücksetzen?", vbYesNo, "Sicherheitsrückfrage")
 If erg <> vbYes Then Exit Function
 archiviert = Now()
 On Error GoTo fehler
' DoCmd.Close acTable, "fuerDiagExp"
' SET rfDEA = TabÖff("fuerDiagExpArchiv", "ID")
' SET rfDE = TabÖff("fuerDiagExp")
' Call rfDEA.Open("fuerdiagexparchiv", DBCn, adOpenDynamic, adLockOptimistic)
 myFrag rfDEA, "SELECT * FROM fuerdiagexparchiv"
' Call rfDE.Open("fuerdiagexp", DBCn, adOpenDynamic, adLockOptimistic)
 myFrag rfDE, "SELECT * FROM fuerdiagexp"
 If Not rfDE.BOF Then
  Do While Not rfDE.EOF
   rfDEA.AddNew
'   ON Error Resume Next
   For Each F In rfDEA.Fields
    If F.name <> "ID" And F.name <> "archiviert" Then
     F.Value = rfDE.Fields(F.name).Value
    End If
   Next F
   rfDEA!archiviert = archiviert
   On Error GoTo rfdeaFehler
   rfDEA.Update
   On Error GoTo fehler
   rfDE.Delete
   rfDE.MoveNext
  Loop
 End If
 If Not frm Is Nothing Then Call knöpfeanpassen(frm)
 Exit Function
rfdeaFehler:
Dim tonRunde%
For tonRunde = 1 To 10
 Call Sound(WinDir + "\media\Windows XP-Standard.wav")
Next tonRunde
 If Err.Number = 3022 Then
  Resume Next
 Else
  MsgBox "Resume in do_Diagnosen_Reset: Fehler beim Aktualisieren von fuerDiagExpArchiv!" & vbCrLf & "Err.Nr:" & Err.Number & vbCrLf & "Err.Description: " & Err.Description
'  Stop
  Resume
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Diagnosen_Reset/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Diagnosen_Reset

#If False Then
Function ÖffneAnamnesebogen()
 On Error GoTo fehler
 Call ÖffneFormular("Anamnesebogen")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ÖffneAnamnesebogen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ÖffneAnamnesebogen
#End If

#If False Then
Function ÖffneFremdlabor()
 On Error GoTo fehler
 Call ÖffneFormular("LaborDokumente")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ÖffneFremdlabor/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ÖffneFremdlabor()
#End If

#If False Then
' wird in und für Commandbars benötigt
Function ÖffneFormular(FoName$)
 Dim Tbl
 On Error GoTo fehler
 Call dtbInit
 On Error Resume Next
 For Each Tbl In Dtb.TableDefs
  DoCmd.Close acTable, Tbl.name, acSaveYes
 Next
 On Error GoTo fehler
 DoCmd.OpenForm FoName$ '"Anamnesebogen nach Name" dtb.Containers(2).Documents(0).Name
 DoCmd.Maximize
 On Error Resume Next
 Forms!Anamnesebogen.Recordset.FindFirst "ID = " + CStr(AktID) ' Wieder alten Datensatz wählen
 On Error GoTo fehler
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ÖffneFormular/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ÖffneFormular
#End If

#If False Then
' wird in und für Commandbars benötigt
Function ÖffneTabelle()
' Dim frm
 On Error Resume Next
 Call AnbogVar(True)
 AktID = Forms(Anmnb)(ABPat_ID)
 DoCmd.Close acForm, Anmnb, acSaveYes ' "Anamnesebogen nach Name"
 On Error GoTo fehler
 DoCmd.OpenTable Dtb.TableDefs!Anamnesebogen.name, acViewDesign, acEdit
 '(dtb.TableDefs!Anamnesebogen!Vorname).SetFocus
' DoCmd.GoToControl "Vorname"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ÖffneTabelle/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ÖffneTabelle
#End If

' wird in und für Commandbars benötigt
Function AnpassForm()
 Dim cb, i%
 Dim PatName$
 On Error GoTo fehler
 Err.Clear
 On Error Resume Next
 Call AnbogVar
 PatName = GesNamFn(Forms(Anmnbi).anaRS) ' Forms!anamnesebogen!Nachname & ", " & Forms!anamnesebogen.Vorname
 On Error GoTo fehler
 If Err.Number > 0 Then PatName = vNS
 Err.Clear
 With CommandBars!Anamneseblatt.Controls(1)
 For i = 1 To .Controls.COUNT
'  Debug.Print .Controls(i).OnAction
  If .Controls(i).OnAction = "wausgeb" Then
   .Controls(i).Caption = "&Werte für " + PatName + " ausgeben" ' wausgeb
   Exit For
  End If
 Next i
 End With
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in AnpassForm/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AnpassForm()

#If False Then
' Toolprogramm, um Allowzerolength zu korrigieren, wird nirgends aufgerufen
Function tabAnp()
Dim Tabl As DAO.tabledef, i%  ', fld AS Object
On Error GoTo fehler
Set Tabl = Dtb.TableDefs("Anamnesebogen")
For i = 1 To Dtb.TableDefs(1).Fields.COUNT
 With Dtb.TableDefs(1).Fields(i)
 If Dtb.TableDefs(1).Fields(i).Type = 10 Then
  If Not Dtb.TableDefs(1).Fields(i).AllowZeroLength Then
   Dtb.TableDefs(1).Fields(i).AllowZeroLength = True
  End If
 End If
 End With
Next
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in tabAnp/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' tabanp()
#End If

' aufgerufen in Datenbank_Aufruf_Click
Function do_Datenbank_Aufruf_Click(frm As Form)
 On Error GoTo fehler
 If QMdbAkt = vNS Then QMdbAkt = LokPfad(QmdB)
 Call Shell(syscmd(9) + "msaccess.exe " & "" & QMdbAkt + vNS, vbNormalFocus)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Datenbank_Aufruf_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Datenbank_Aufruf_Click(frm AS Form)

' wird aus dem Formularcode aufgerufen
Function do_Ausgabe_Click(frm As Form)
'On Error GoTo Err_Ausgabe_Click
 On Error GoTo fehler
 Call GibwerteAus '(frm!Pat_ID)

'    Call Shell("NOTEPAD.EXE """ & uverz & "Ergeb.txt""", 1)

 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Ausgabe_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Ausgabe_Click(frm As Form)

#If mitab Then
' wird aus dem Formularcode aufgerufen
Function do_Datenquelle_Change(frm As Form)
 Dim Nachname$, Vorname$, GebDat$, altpat_id
 On Error GoTo fehler
 Select Case frm.name
  Case Anmnbi ' "Anamnesebogen"
   Nachname = vNS
   On Error Resume Next
   Nachname = frm.Recordset!Nachname
   Vorname = frm.Recordset!Vorname
   GebDat = frm.Recordset!GebDat
   frm.RecordSource = frm.Datenquelle
   On Error GoTo fehler
   If Nachname <> vNS Then
    On Error Resume Next
    frm.Recordset.FindFirst "NachName = '" + Nachname + "' AND vorname = '" + Vorname + "' AND format$(gebdat,""dd/mm/yyyy"") = '" + GebDat + "'"
    On Error GoTo fehler
   End If
  Case "LaborDokumente"
   altpat_id = vNS
   On Error Resume Next
   altpat_id = CStr(frm.Recordset!Pat_ID)
   frm.RecordSource = frm.Datenquelle
   On Error GoTo fehler
   If altpat_id <> vNS Then
    On Error Resume Next
    frm.Recordset.FindFirst "pat_id = " & altpat_id
    On Error GoTo fehler
   End If
 End Select
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Datenquelle_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Datenquelle_Change(frm As Form)

Function DqC_U(frm As Form)
 On Error GoTo fehler
 frm.Datenquelle = "Anamnesebogen Unausgefüllte"
Call do_Datenquelle_Change(frm)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DqC_U/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DqC_U(frm AS Form)

Function DqC_UNZ(frm As Form)
 On Error GoTo fehler
 frm.Datenquelle = "Anamnesebogen Unausgefüllte nach Zeit"
 Call do_Datenquelle_Change(frm)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DqC_UNZ/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DqC_UNZ(frm AS Form)

Function DqC_A(frm As Form)
 On Error GoTo fehler
 frm.Datenquelle = "Anamnesebogen alle Datensätze"
 Call do_Datenquelle_Change(frm)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DqC_A/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DqC_A(frm AS Form)
#End If

#If mitab Then
' wird aus dem Formularcode aufgerufen
Function do_Form_Open(Cancel%, frm As Form)
#If False Then
 Dim rs As DAO.Recordset
 Dim q As DAO.QueryDef
 Dim geeignet%
 On Error GoTo fehler
 If QMdbAkt = vNS Then QMdbAkt = LokPfad(QmdB)
 Select Case frm.name
  Case "LaborDokumente eP"
   frm.RecordSource = "SELECT * FROM `" + QMdbAkt + "`.`" + frm.name + "`"
   Exit Function
  Case "Checklisten"
   frm.RecordSource = "SELECT * FROM `" + QMdbAkt + "`.`" + frm.name + "`"
   Exit Function
  Case "Augenbefunde"
   frm.RecordSource = "SELECT * FROM `" + QMdbAkt + "`.`" + frm.name + "`"
  Case "Fremdlabor altes Konzept"
   frm.RecordSource = "SELECT * FROM `" + QMdbAkt + "`.`" + "Fremdlabor" + "`"
  Case "LaborDokumente"
   frm.RecordSource = "SELECT * FROM `" + QMdbAkt + "`.`" + "LaborDokumente fAB" + "`"
 End Select
 
 ' aus do_Fremdlabor_Form_Open
 Select Case frm.name
  Case "Augenbefunde", "Fremdlabor altes Konzept"
   frm!FDatenquelle.RowSource = vNS
   For Each q In Dtb.QueryDefs
    If InStr(q.name, frm.name) = 1 Then
     frm!FDatenquelle.RowSource = frm!FDatenquelle.RowSource + q.name + ";"
    End If
   Next q
   Exit Function
 End Select
 ' hier sind also noch "Anamnesebogen", "Checklisten","LaborDokumente"
 frm.Datenquelle.RowSource = vNS
 Call dtbInit
 For Each q In Dtb.QueryDefs
  geeignet = 0
  Select Case frm.name
   Case Anmnbi ' "Anamnesebogen"
     If InStr(q.name, "Anamnese") = 1 Then geeignet = 1
     frm![Datenbank-Aufruf].Caption = QMdbAkt
   Case "LaborDokumente"
     If InStr(q.name, "LaborDokumente") = 1 Then geeignet = -1
  End Select
  If geeignet Then
   frm.Datenquelle.RowSource = frm.Datenquelle.RowSource + q.name + ";"
  End If
 Next q
' For Each q In currentDB.QueryDefs
'  geeignet = 0
'  SELECT CASE frm.name
'   Case "Anamnesebogen"
'     IF InStr(q.name, "Anamnese") = 1 THEN geeignet = 1
'     frm.`datenbank-Aufruf`.Caption = qmdbakt
'   Case "LaborDokumente"
'     IF InStr(q.name, "LaborDokumente") = 1 THEN geeignet = -1
'  END SELECT
'  IF geeignet AND instrb(frm.Datenquelle.RowSource, q.name) = 0 THEN
'   frm.Datenquelle.RowSource = frm.Datenquelle.RowSource + q.name + ";"
'  END IF
' Next q

' IF instrb(frm.Name, "Anamnesebogen") > 0 AND rfDE Is Nothing THEN
'  SET rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
' END IF

 'If rESt Is Nothing THEN
' SET rest = TabÖff("Einstellungen", "Formular")
' ON Error Resume Next
' rest.Seek "=", frm.Name
' IF Not rest.NoMatch THEN
'  frm.Datenquelle = rest!`Abfrage für Formular`
'  frm.RecordSource = frm.Datenquelle
''  frm.Recordset.FindFirst "pat_ID = " + 97 'CStr(rESt!`id für formular`) ' geht irgendwie nicht (mehr?)
'  DoCmd.GoToRecord acDataForm, "Anamnesebogen", acGoTo, rest!DatensatzNr
' END IF
' ON Error GoTo fehler
' rest.Close
#End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Form_Open/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Form_Open(Cancel%, frm AS Form)
#End If

#If mitab Then
' wird aus dem Formularcode aufgerufen
Function do_Form_Close(frm As Form)
#If False Then
' IF rESt Is Nothing THEN
 On Error GoTo fehler
 Set rest = TabÖff("Einstellungen", "Formular")
 If Not rest Is Nothing Then
  rest.Seek "=", frm.name
  If rest.Nomatch Then
   rest.AddNew
   rest!Formular = frm.name
  Else
   rest.Edit
  End If
  rest!`Abfrage für Formular` = frm.Datenquelle
  On Error Resume Next
  rest!`ID für Formular` = frm.Recordset!Pat_id
  rest!`DatensatzNr` = frm.Recordset.AbsolutePosition + 1
  rest.Update
  On Error GoTo fehler
  rest.Close
 End If
 'Call übrigeRecordsetsSchließen
 #End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Form_Close/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Form_Close(frm AS Form)
#End If

'Function übrigeRecordsetsSchließen()
' ON Error GoTo fehler
' IF Not rfDE Is Nothing THEN
'   rfDE.Close
'   SET rfDE = Nothing
' END IF
' IF Not aNaüm Is Nothing THEN
'   aNaüm.Close
'   SET aNaüm = Nothing
' END IF
' IF Not rLaU Is Nothing THEN
'  rLaU.Close
'  SET rLaU = Nothing
' END IF
' IF Not rDokab Is Nothing THEN
'   rDokab.Close
'   SET rDokab = Nothing
' END IF
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIF(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in übrigeRecordsetsSchließen/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End Function
Function do_form_beforeUpdate(frm)
 On Error GoTo fehler
 If Not frm.Recordset.BOF Then
  On Error Resume Next
  altpat_id = frm.Recordset!Pat_ID
  On Error GoTo fehler
  ReDim altDS(frm.Recordset.Fields.COUNT - 1)
  On Error Resume Next
  For DSNr = 0 To frm.Recordset.Fields.COUNT - 1
   altDS(DSNr) = frm.Recordset.Fields(DSNr)
  Next DSNr
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_form_beforeUpdate/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_form_beforeUpdate(frm)

' wird aus dem Formularcode aufgerufen
Function do_Form_AfterUpdate(frm)
 Dim obGleich As Boolean
' IF altID = frm!Pat_id THEN
 On Error GoTo fehler
  obGleich = True
   On Error Resume Next
   If frm.Recordset!Pat_ID = altpat_id Then
    For DSNr = 0 To frm.Recordset.Fields.COUNT - 1
'     IF dsNr = 17 THEN
'      Debug.Print "frm.Recordset.Fields(dsNr):", frm.Recordset.Fields(dsNr)
'      Debug.Print "altDS(dsNr):", altDS(dsNr)
'     END IF
     If frm.Recordset.Fields(DSNr) <> altDS(DSNr) Or _
     (IsNull(frm.Recordset.Fields(DSNr)) And (Not IsNull(altDS(DSNr)))) Or _
     (Not (IsNull(frm.Recordset.Fields(DSNr))) And IsNull(altDS(DSNr))) Then
      If frm.Recordset.Fields(DSNr).name <> "letzte Änderung" Then
       obGleich = False
      End If
      Exit For
     End If
    Next DSNr
    On Error GoTo fehler
'   END IF
   If Not obGleich Then
    Call RRParseF(frm!Pat_ID)
    frm![letzte Änderung] = Now
   End If
  End If
' END IF
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Form_AfterUpdate/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'do_Form_AfterUpdate(frm)

#If zutesten Then
Function formtest()
 Dim i%
 On Error GoTo fehler
 For i = 1 To Forms.COUNT
  Debug.Print Forms(i).name
 Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in formtest/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' formtest()
#End If

' wird aus dem Formularcode aufgerufen
Function do_Form_Current_Medarten(frm As Medarten)

End Function ' do_Form_Current_Medarten(frm As Medarten)
#If mitab Then

' aufgerufen in anaRS_MoveComplete
Function do_Form_Current_AnBog(frm As AnBog)
 Dim rsNa As New ADODB.Recordset
 Dim farbe&, i&, tStr$
' Static rsAnam As Adodb.Recordset
' SET rsAnamlt = frm.Recordset
' altPat_id = frm.Recordset!Pat_id
' ReDim altDS(frm.Recordset.Fields.Count - 1)
' For dsNr = 0 To frm.Recordset.Fields.Count - 1
'  altDS(dsNr) = frm.Recordset.Fields(dsNr)
' Next dsNr
' Debug.Print "Form_Current 1:" + format$(Now, "hh:mm:ss")
 Dim ctrl As Control

' Diagnosen
 Dim KZahlNeu%, pnpflNeu%, anpflNeu%, retflNeu%, katflNeu%, bgwflNeu%, _
 angflNeu%, dfsflNeu%, fieflNeu%, oblhNeu%
 On Error GoTo fehler
 If frm.obStumm Then Exit Function
#If debu <> 0 Then
       Open Environ("allusersprofile") & "\Application Data\chron.txt" For Output As #313
       Print #313, Now() & " " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
       dnr = 0
       T0 = Timer
       Takt = T0
#End If
 frm.obStumm = True
 KZahlNeu = 0
 pnpflNeu = 0
 anpflNeu = 0
 retflNeu = 0
 katflNeu = 0
 bgwflNeu = 0
 oblhNeu = 0
 angflNeu = 0
 dfsflNeu = 0
 fieflNeu = 0
 
 On Error Resume Next
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", vor Markierfarbe)"
#End If
 For Each ctrl In frm.Controls
  If ctrl.BackColor = MarkierFarbe Then ctrl.BackColor = -2147483643
 Next ctrl
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", nach Markierfarbe)"
#End If
 On Error GoTo fehler
' dtbInit
' Set rsNa = Nothing
'  SET rsNa = TabÖff("Anamnesebogen", "Pat_id")
 Dim altpat_id&
 altpat_id = Pat_ID
 If frm.anaRS.EOF Or frm.anaRS.BOF Then
'  Pat_id = 0
  Pat_ID = myEFrag("select pat_id FROM (" & frm.anaRS.source & ") i ORDER BY vorgestellt " & IIf(frm.anaRS.BOF, "DESC", "")).Fields(0)
 ElseIf IsNull(frm.anaRS!Pat_ID) Then
  frm.anaRS.MoveFirst
  Pat_ID = frm.anaRS!Pat_ID
'  Pat_id = myEFrag("SELECT MAX(pat_id) FROM (" & frm.anaRS.source & ") i").Fields(0)
 Else
  Pat_ID = frm.anaRS!Pat_ID
 End If
 If Pat_ID = altpat_id Then
  Tüt 2740, 40
'  TütSoundkarte
  GoTo kurzvorEnde
 End If
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 Set rsNa = Nothing
 myFrag rsNa, "SELECT * FROM `anamnesebogen` WHERE pat_id = " & Pat_ID
 If rsNa.BOF Then
  frm.obStumm = False
  Exit Function
 End If
 On Error Resume Next
 syscmd 4, "Beginn Formularvorbereitung " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", vor Diagnosen)"
#End If
 Err.Clear
' tStr = Pat_id
' IF Err.Number <> 0 THEN
'  MsgBox "Falsche Abfrage: " + frm!Datenquelle
'  MsgBox (currentDB.QueryDefs(frm!Datenquelle).sql)
'  MsgBox "Ein Feld muß mit 'pat_id' ohne Vorsätze ansprechbar sein"
' END IF
' nzl = vbcr + vblf
' rsNa.Seek "=", pat_id
 If rsNa.BOF Then
  frm.vTextB(147) = vNS ' Diagnosen
 Else
'  frm.vTextB(147) = replace$(replace$(rsNa!Diagnosen, vbverticaltab, vbcr + vblf), vbtab, chr$(32))
  Dim DiagTab() As CString, zwiDiag$
  ' frm.vTextB(5) = Diabetes seit
  zwiDiag = REPLACE$(REPLACE$(DiagString(CStr(Pat_ID), DiagTab, , , frm.vTextB(5)), vbVerticalTab, vbCrLf), vbTab, Chr$(32))
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", in Diagnosen)"
#End If
   frm.vTextB(147) = zwiDiag
 End If
 frm.vTextB(181) = lebe(Pat_ID) ' leBeh
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", nach Diagnosen)"
#End If
 frm.vTextB(188) = frm.anaRS!vorET
' IF (ISNULL(frm.anaRS!Ther1) OR frm.anaRS!Ther1 = vns) AND frm.anaRS!Diabetestyp <> "-" THEN
'  frm.vTextB(166) = TherUmw(TherArt(Pat_id, -1))
'  Debug.Print "frm.anaRS!Ther1:", frm.anaRS!Ther1
'  Dim ta$ ' 7.9.09
'  ta = TherUmw(TherArt(Pat_id, -1))
'  IF frm.anaRS!Ther1 <> ta THEN
' frm.obstumm = False   ' 7.9.09
' Exit FUNCTION          ' 7.9.09
'   frm.vTextB(166) = ta
'   Debug.Print frm.vTextB(166) & " -> " & ta
'   Stop
'  END IF
'  frm.anaRS.Save
' END IF
' IF (ISNULL(frm.anaRS!Ther1) OR frm.anaRS!Ther1 = vns) AND frm.anaRS!Diabetestyp <> "-" THEN
'  IF frm.anaRS!Ther1 <> TherUmw(TherArt(Pat_id, -1)) THEN frm.vTextB(166) = TherUmw(TherArt(Pat_id, -1))
' END IF
' IF (ISNULL(frm.anaRS!TherAkt) OR frm.anaRS!TherAkt = vns) AND frm.anaRS!Diabetestyp <> "-" THEN
'  IF frm.anaRS!TherAkt <> TherUmw(TherArt(Pat_id, 0)) THEN frm.vTextB(173) = TherUmw(TherArt(Pat_id, 0))
' END IF

' IF rafDE Is Nothing THEN Call rafDE.Open("SELECT * FROM `fuerdiagexp`", DBCn, adOpenDynamic, adLockReadOnly) 'Set rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
' If HArst Is Nothing Then Call HArst.Open("SELECT * FROM `aktlue`", DBCn, adOpenDynamic, adLockReadOnly) 'Set HArst = TabÖff("Hausaerzte", "KVNr") ' isempty(harst), falls über dim harst statt dim harst AS dao.recordset definiert
 Static HArst As New ADODB.Recordset
 If HArst Is Nothing Then myFrag HArst, "SELECT * FROM `aktlue`"
 On Error Resume Next
 
 Dim obHAimDMP%
 obHAimDMP = 0
 frm.vTextB(148) = vNS ' HA
 frm.vTextB(176) = vNS ' HAFax
 frm.vTextB(169) = vNS ' HAimDMP
' Dim HaNr1$, HANr2$
 frm.vComboB(1) = vNS ' HA_Auswahl
 frm.vComboB(2) = vNS ' HA_auswahl2
 Dim HaErg$(), wHAErg$
 wHAErg = Üw12(Pat_ID, HaErg)
' frm.vTextB(183) = wHAErg ' 7.9.09
' frm.vTextB(183) = Üw12(Pat_id, HaErg) ' HANr, neues Feld ' 7.9.09
 If LenB(frm.vTextB(183)) <> 0 Then If UBound(HaErg, 2) > 0 Then frm.vTextB(175) = HaErg(0, 1) ' HANr2
 ReDim HANrBf(2)
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 HANrBf(1) = wHAErg ' frm.vTextB(183) ' HANr <- Ersatz schaffen
 HANrBf(2) = frm.vTextB(175) ' HANr2
 frm.vCheckb(15) = (HaErg(0, 0) = vNS) ' HA1nicht
 If frm.vCheckb(15) <> 0 Then
  If UBound(HaErg, 2) > 0 Then
   frm.vCheckb(16) = (HaErg(0, 1) = vNS) ' HA2nicht
  End If
 End If
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 If Not IsNull(wHAErg) And LenB(wHAErg) <> 0 Then ' ISNULL(frm.vTextB(183)) AND frm.vTextB(183) <> vns THEN ' HANr
'  HArst.FindFirst "ltrim$(KVNr) = ltrim$('" + frm.HANr + "')"

' folgendes geht schlecht in ein Unterprogramm
   Dim teile$()
   teile = Split(HANrBf(1), "_")
       Dim gefunden%
       gefunden = 0
       If UBound(teile) >= 2 Then
'       rHa.Seek "=", teile(0), trim$(teile(1)), trim$(teile(2))
        Set HArst = Nothing
        myFrag HArst, "SELECT * FROM `aktlue` WHERE kvnro = '" & teile(0) & "' AND nameo = '" & teile(1) & "' AND vno = '" & teile(2) & "'"
        If Not HArst.BOF Then gefunden = True
       End If
       If UBound(teile) = 1 Or (UBound(teile) > 1 And Not gefunden) Then
        gefunden = 0
'       rHa.Seek "=", teile(0), trim$(teile(1))
        Set HArst = Nothing
        myFrag HArst, "SELECT * FROM `aktlue` WHERE kvnro = '" & teile(0) & "' AND nameo = '" & teile(1) & "'"
        If Not HArst.BOF Then gefunden = True
       End If
       If UBound(teile) = 0 Or (UBound(teile) > 0 And Not gefunden) Then
        gefunden = 0
 '      rHa.Seek "=", teile(0)
        Set HArst = Nothing
        myFrag HArst, "SELECT * FROM `aktlue` WHERE kvnro = '" & teile(0) & "'"
       End If
      
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
   If HArst.BOF Or IsNull(HArst!name) Then
'    IF ISNULL(HArst!Name) THEN
'     HArst.Delete
'    END IF
    Select Case UBound(teile)
     Case 2: Call HAlokal(HArst, teile(0), teile(1), teile(2))
     Case 1: Call HAlokal(HArst, teile(0), teile(1))
     Case 0: Call HAlokal(HArst, teile(0))
    End Select
   End If
   If Not HArst.BOF And Not UBound(teile) = -1 Then ' wenn gar nicht gesucht wurde
    frm.vTextB(148) = IIf(IsNull(HArst!name), vNS, HArst!name) + ",T:" + IIf(IsNull(HArst!telefon), vNS, HArst!telefon) & ", " & IIf(IsNull(HArst!strasse), vNS, HArst!strasse) & " " & IIf(IsNull(HArst!plz), vNS, HArst!plz) & " " & IIf(IsNull(HArst!ort), vNS, HArst!ort) + ", Z:" + IIf(IsNull(HArst!zulg), vNS, HArst!zulg)
    frm.vTextB(176) = IIf(IsNull(HArst!fax), vNS, HArst!fax)
    obHAimDMP = fobHAimDMP(frm.vTextB(183)) ' HANr
   End If
   HArst.Find "kvNro=" & LTrim$(frm.vTextB(183))    ' HANr
   frm.vTextB(169) = IIf(obHAimDMP, "Hausarzt im DMP", "Hausarzt nicht im DMP") ' HAimDMP
 End If
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
' IF rFa Is Nothing THEN SET rFa = TabÖff("faelle", "aktF")
' rFa.Seek "=", pat_id
' Set raFa = Nothing
' Call raFa.Open("SELECT * FROM `faelle` WHERE pat_id = " & frm!Faelle, DBCn, adOpenDynamic, adLockReadOnly)
' frm.anaRS!SchGr = raFa!SchGr
 Select Case frm.anaRS!Versicherungsart ' Scheingruppe
  Case "00"
   farbe = 16711808 ' lila
  Case "24"
   farbe = 65280 ' grün
  Case "41"
   farbe = 12632256 ' hellgrau
  Case "90"
   farbe = 65535 ' gelb
  Case Else
   farbe = -2147483633  ' weiß
 End Select
 frm.MA0Zahl.BackColor = IIf(frm.anaRS!obtk, &HFF0000, -2147483633)
 frm.MA1Zahl.BackColor = IIf(frm.anaRS!obgs, 65535, -2147483633)
 frm.MA2Zahl.BackColor = IIf(frm.anaRS!obdw, 65535, -2147483633)
 frm.MA3Zahl.BackColor = IIf(frm.anaRS!obah, 65535, -2147483633)
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 frm.vTextB(160).BackColor = farbe ' NachName
 frm.vTextB(2).BackColor = farbe ' Vorname
 frm.vTextB(152).BackColor = farbe ' Titel
 frm.vTextB(153).BackColor = farbe ' Anrede
 
 If frm.anaRS!Größe = 0 Or frm.anaRS!Gewicht = 0 Then
   farbe = -2147483633 ' weiß
 Else
   On Error Resume Next
   Dim bmi!, Gewicht!, Größe!
   Dim runde%
   runde = 0
   If IsNumeric(frm.vTextB(142)) Then bmi = frm.vTextB(142)
   If IsNumeric(frm.vTextB(11)) Then Gewicht = frm.vTextB(11)
   If IsNumeric(frm.vTextB(10)) Then Größe = frm.vTextB(10)
   Do
    runde = runde + 1
    If bmi = 0 Then Exit Do ' BMI
    If bmi > 8 Then Exit Do ' BMI
    bmi = bmi * 10 ' BMI
    If Gewicht < 10 Then
     Gewicht = Gewicht * 10
    End If
   Loop Until runde >= 10
   runde = 0
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
   Do
    If bmi = 0 Then Exit Do ' BMI
    If bmi < 80 Then Exit Do ' BMI
'    frm.vtextb(142) = frm.vtextb(142) * 0.1 => geht nicht ,da Berechnungsfeld
    If Gewicht > 250 Then ' Gewicht     <- Ersatz schaffen
       Gewicht = Gewicht * 0.1 ' Gewicht
    ElseIf Größe < 3 Then ' Größe     <- Ersatz schaffen
      Größe = Größe * 10 ' Größe
    End If
    runde = runde + 1
   Loop Until runde >= 10
   Select Case bmi ' BMI
    Case Is < 20
     farbe = 15263952
    Case Is < 25
     farbe = 13619100
    Case Is < 30
     farbe = 11369810
    Case Is < 35
     farbe = 10314363
    Case Is < 40
     farbe = 8673192
    Case Else
     farbe = 5387467
   End Select
   On Error GoTo fehler
 End If
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 frm.vTextB(142).BackColor = farbe ' BMI
 frm.vTextB(10).BackColor = farbe ' Größe
 frm.vTextB(11).BackColor = farbe ' Gewicht
 frm.vTextB(12).BackColor = farbe ' Tendenz
 
 syscmd 4, "Formularvorbereitung 1 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 If IsNumeric(frm.anaRS![letztes HbA1c]) Then ' letztes HbA1c
  Select Case Val(frm.anaRS![letztes HbA1c]) ' letztes HbA1c
   Case 0
    frm.vTextB(43).BackColor = -2147483643 ' letztes HbA1c
   Case Is < 6
    frm.vTextB(43).BackColor = 15263952 ' letztes HbA1c
   Case Is <= 6.5
    frm.vTextB(43).BackColor = 13619100 ' letztes HbA1c
   Case Is < 7
    frm.vTextB(43).BackColor = 11369810 ' letztes HbA1c
   Case Is < 8
    frm.vTextB(43).BackColor = 10314363 ' letztes HbA1c
   Case Is < 9
    frm.vTextB(43).BackColor = 8673192 ' letztes HbA1c
   Case Else
    frm.vTextB(43).BackColor = 5387467 ' letztes HbA1c
  End Select
 Else
   frm.vTextB(43).BackColor = -2147483643 ' weiß  ' letztes HbA1c
 End If
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
' Dim lsql$
' lsql = "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_id & " UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_id & ") i GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb"
 Set raLau = Nothing
' raLau.Open "SELECT * FROM (" & lsql & ") AS innen WHERE abkü LIKE '%BNP%' ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
 Set raLau = hollabor(Pat_ID, "BNP")
 If Not raLau.EOF Then
  frm.ntPro = raLau!Wert
 Else
  frm.ntPro = vNS
 End If
 Set raLau = Nothing
 
' raLau.Open "SELECT * FROM (" & lsql & ") AS innen WHERE abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBU') AND (abkü <> 'ALBU' OR wert LIKE '%<%') ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
#If False Then
' frm.AlbCre = vNS
' SET raLau = LabEPat(AlbCre, Pat_id)
' Do While Not raLau.EOF
'  frm.AlbCre = frm.AlbCre & ROUND(REPLACE(raLau!wert, ".", ",")) & "/"
'  raLau.Move 1
' Loop
#ElseIf False Then
' myFrag raLau, "SELECT GROUP_CONCAT(CAST(wert AS char) SEPARATOR '-') g " & vbCrLf & _
'             "FROM (SELECT DATE(zeitpunkt) zp, ROUND(IF(ISNULL(wert),IF(ISNULL(kommentar),'',kommentar),wert)) Wert " & vbCrLf & _
'             "FROM `laborneu` ln LEFT JOIN laborkommentar lk ON ln.kommentarvw = lk.kommentarvw " & vbCrLf & _
'             "WHERE ((abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND einheit LIKE 'mg/g %') OR (abkü IN ('ALBU','ALBUMU') AND (einheit = 'mg/l' OR einheit = ''))) AND pat_id = " & Pat_id & " " & vbCrLf & _
'             "UNION " & vbCrLf & _
'             "SELECT DATE(u.eingang) zp, ROUND(IF(ISNULL(w.wert),IF(ISNULL(t.text),'',t.text),w.wert)) Wert " & vbCrLf & _
'             "FROM `laboryus` u " & vbCrLf & _
'             "LEFT JOIN laborywert w ON u.id = w.usid LEFT JOIN laboryhinw t ON w.kommid = t.id " & vbCrLf & _
'             "WHERE ((abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND einheit LIKE 'mg/g %') OR (abkü IN ('ALBU','ALBUMU') AND (einheit = 'mg/l' OR einheit = ''))) AND pat_id = " & Pat_id & " " & vbCrLf & _
'             "GROUP BY zp ORDER BY zp DESC) i;"
' IF NOT ISNULL(raLau!g) THEN
'  frm.AlbCre = raLau!g
' Else
'  frm.AlbCre = vNS
' END IF
#Else
 Dim Labs As labtyp
 frm.AlbCre = vNS
 alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
 Do While True
  Labs = LabPat(LA_AlbCre, Pat_ID, True)
  If Labs.Zp = -1 Or Labs.Zp = #12/29/1899# Then Exit Do
  If frm.AlbCre <> vNS Then frm.AlbCre = frm.AlbCre & "-"
  If IsNumeric(Labs.WertSg) Then frm.AlbCre = frm.AlbCre & Round(Labs.WertSg)
 Loop
#End If
 
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 If IsNull(frm.anaRS!Diabetestyp) Or frm.anaRS!Diabetestyp = vNS Then ' Diabetestyp
  farbe = 16777215 ' weiß
 Else
  Select Case frm.anaRS!Diabetestyp ' Diabetestyp
   Case "1"
    farbe = 255 'rot
   Case "2"
    farbe = 33023  ' orange
   Case "S", "s"
    farbe = 12615935 ' rosa
'   Case vns
'    farbe = 16777215 ' weiß
   Case "-"
    farbe = 12632256 ' hellgrau
   Case Else
    farbe = 65280 ' grün
  End Select
 End If
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 frm.vTextB(4).BackColor = farbe ' Diabetestyp
 frm.vTextB(5).BackColor = farbe ' `Diabetes seit`
 frm.vTextB(6).BackColor = farbe ' `Tabletten seit`
 frm.vTextB(7).BackColor = farbe ' `Insulin seit`
' Schwangerschaftsfelder ggf. sperren
 Dim SAkt%
 SAkt = -1
 If Not IsNull(frm.anaRS!anrede) Then
  If Not IsNull(frm.anaRS!GebDat) Then
   If frm.anaRS!anrede = "Herr" Or Date - CDate(frm.anaRS!GebDat) > 55 * 365 Then
    SAkt = 0
   End If
  End If
 End If
 If SAkt Then
  Call Acti(frm.vTextB(63)) ' frm.anaRS!Schwanger
  Call Acti(frm.vTextB(64)) ' `Schwanger seit`
 Else
  Call inActi(frm.vTextB(63)) ' frm.anaRS!Schwanger
  Call inActi(frm.vTextB(64)) ' `Schwanger seit`
 End If
 If frm.anaRS!Ther1 = "OAD" Then
  Call inActi(frm.vTextB(37)) ' `Spritzstelle früh (B, O, A)`
  Call inActi(frm.vTextB(38))   ' `Spritzstelle mittags`
  Call inActi(frm.vTextB(39)) ' `Spritzstelle abends`
  Call inActi(frm.vTextB(40)) ' `Spritzstelle nachts`
 Else
  Call Acti(frm.vTextB(37)) ' `Spritzstelle früh (B, O, A)`
  Call Acti(frm.vTextB(38))   ' `Spritzstelle mittags`
  Call Acti(frm.vTextB(39)) ' `Spritzstelle abends`
  Call Acti(frm.vTextB(40)) ' `Spritzstelle nachts`
 End If
 If frm.anaRS![BZMessungen selbst] = "n" Or frm.anaRS![BZMessungen selbst] = "-" Then
  Call inActi(frm.vTextB(158))   ' Gerät
  Call inActi(frm.vTextB(49)) ' Aufschreiben
 Else
  Call Acti(frm.vTextB(158))   ' Gerät
  Call Acti(frm.vTextB(49)) ' Aufschreiben
 End If
 If frm.anaRS!Bluthochdruck = "n" Or frm.anaRS!Bluthochdruck = "-" Then
  Call inActi(frm.vTextB(59)) ' BHD seit
  Call inActi(frm.vTextB(60)) ' BHD beh mit
 Else
  Call Acti(frm.vTextB(59)) ' BHD seit
  Call Acti(frm.vTextB(60)) ' BHD beh mit
 End If
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 If frm.anaRS!Herzkrankheit = "n" Or frm.anaRS!Herzkrankheit = "-" Then
  Call inActi(frm.vTextB(75)) ' `Angina pectoris`
  Call inActi(frm.vTextB(76)) ' `Herzinfarkt`
  Call inActi(frm.vTextB(77)) ' `Herzinfarkt wann`
  Call inActi(frm.vTextB(78)) ' `PTCA oder Stent`
  Call inActi(frm.vCheckb(3)) ' [Bypaß kardial]
  Call inActi(frm.vTextB(79)) ' `Bypass wann`
  Call inActi(frm.vTextB(80)) ' `Herzschwäche`
  Call inActi(frm.vTextB(81)) ' [Herzkrankheit Beschreibung]
 Else
  Call Acti(frm.vTextB(75)) ' `Angina pectoris`
  Call Acti(frm.vTextB(76)) ' `Herzinfarkt`
  Call Acti(frm.vTextB(77)) ' `Herzinfarkt wann`
  Call Acti(frm.vTextB(78)) ' `PTCA oder Stent`
  Call Acti(frm.vCheckb(3)) ' [Bypaß kardial]
  Call Acti(frm.vTextB(79)) ' `Bypass wann`
  Call Acti(frm.vTextB(80)) ' `Herzschwäche`
  Call Acti(frm.vTextB(81)) ' [Herzkrankheit Beschreibung]
 End If
 If frm.anaRS!Sexualstörung = "n" Or frm.anaRS!Sexualstörung = "-" Then
  Call inActi(frm.vTextB(101)) ' `Sexualstörung seit`
 Else
  Call Acti(frm.vTextB(101)) ' `Sexualstörung seit`
 End If
' Feld mit den Diagnosen der Medikamente füllen
syscmd 4, "Formularvorbereitung 2 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
Dim rsmd As New ADODB.Recordset
Dim hmg$, hypt$, neurp$, autnp$, fetts$, Hsre$, antimyk$, glauk$, cold$, pros$, urä$, hythy$, ostp$, _
    khk$, herzis$, stru$, avk$, pani$, park$, vari$, östr$, antidep$, antidem$, antiep$, antipern$, antiherp$, Antikoag$
    
If Pat_ID <> 0 Then
' SET aNaüm = TabÖff("Anamnesebogen", "Pat_id")
' aNaüm.Seek "=", pat_id
' SET aNaüm = Nothing
' myFrag aNaüm, "SELECT * FROM `anamnesebogen` WHERE pat_id = " & Pat_id
' Vorbelegungen für (zusätzliche) Diagnosen im Feld mdi
 frm.vTextB(168) = vNS ' !MDi
 MDIn = 1
 For i = 0 To UBound(MDIICD)
  MDIICD(i) = vNS
  MDIDiag(i) = vNS
 Next
 Set fm = frm
 Call KrStart
 
 Dim TabakStatus As ZigSt
 TabakStatus = doTabakSt(frm.anaRS!Pat_ID)
 Select Case TabakStatus
  Case 0
  Case 1
   Call KRAdd(frm, frm.anaRS!tabakex, "Nikotinabusus", "F17.1", Zn, "F17.1", , , , "vTextB", 184) ' Tabak
  Case 3
   Call KRAdd(frm, frm.anaRS!tabakex, "Nikotinabusus vor langem", "F17.1", Zn, "F17.1", , , , "vTextB", 184) ' Tabak
  Case 2
   Call KRAdd(frm, frm.anaRS!tabakakt, "Nikotinabusus", "F17.1", gesi, "F17.1", , , , "vTextB", 186) ' "Tabak"
 End Select
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 
 ' andersrum geordnet in EigTest
 'Set md = dtb.OpenRecordset("SELECT Fi.pat_id AS Pat_ID, Fi.zeitpunkt AS Zeitpunkt, Fi.med, `medarten`.HMG, `medarten`.Hypt, `medarten`.Neurp, `medarten`.AutNP, `medarten`.Fetts, `medarten`.Hsre, `medarten`.AntiMyk, `medarten`.Glauk, `medarten`.COLD, `medarten`.Pros, `medarten`.Urä, `medarten`.HyThy, `medarten`.Ostp, `medarten`.KHK, `medarten`.HerzI, `medarten`.Stru, `medarten`.AVK, `medarten`.PanI, `medarten`.Park, `medarten`.Vari, `medarten`.Vari, `medarten`.Östr, `medarten`.AntiDep, `medarten`.AntiDem, FROM `SELECT pat_id, zeitpunkt, (LEFT(feldinh,iIF(instrb(feldinh,"" "")>0,instr(feldinh,"" ""),len(feldinh)))) AS med FROM forminhalt WHERE form_abk = ""mp"" AND feld = ""medikament""`. AS Fi LEFT JOIN `medarten` ON Fi.med = `medarten`.Medikament WHERE fi.pat_id =" + CStr(frm.Pat_ID), dbOpenDynaset)
 sql = "SELECT mp.*, ma.* FROM `medplan` mp LEFT JOIN `medarten` ma ON mp.medanfang = ma.medikament WHERE mp.pat_id =" & Pat_ID
 Set rsmd = Nothing
 myFrag rsmd, sql
' Debug.Print "Form_Current 2:" + format$(Now, "hh:mm:ss")
 If Not rsmd.BOF Then
  Do While Not rsmd.EOF
   Call mdTest(rsmd, "hmg", hmg)
   Call mdTest(rsmd, "hypt", hypt)
   Call mdTest(rsmd, "neurp", neurp)
   Call mdTest(rsmd, "autnp", autnp)
   Call mdTest(rsmd, "fetts", fetts)
   Call mdTest(rsmd, "Hsre", Hsre)
   Call mdTest(rsmd, "antimyk", antimyk)
   Call mdTest(rsmd, "glauk", glauk)
   Call mdTest(rsmd, "cold", cold)
   Call mdTest(rsmd, "pros", pros)
   Call mdTest(rsmd, "urä", urä)
   Call mdTest(rsmd, "hythy", hythy)
   Call mdTest(rsmd, "ostp", ostp)
   Call mdTest(rsmd, "khk", khk)
   Call mdTest(rsmd, "herzi", herzis)
   If InStrB(rsmd!Medikament, "jod") Then ' 16.12.19
    Call mdTest(rsmd, "stru", stru)
   End If
   Call mdTest(rsmd, "avk", avk)
   Call mdTest(rsmd, "pani", pani)
   Call mdTest(rsmd, "park", park)
   Call mdTest(rsmd, "vari", vari)
   Call mdTest(rsmd, "östr", östr)
   Call mdTest(rsmd, "antidep", antidep)
   Call mdTest(rsmd, "antidem", antidem)
   Call mdTest(rsmd, "antiep", antiep)
   Call mdTest(rsmd, "antipern", antipern)
   Call mdTest(rsmd, "antiherp", antiherp)
   Call mdTest(rsmd, "antikoag", Antikoag)
   rsmd.MoveNext
  Loop ' not rsmd.eof
 End If ' not rsmd.bof
' Debug.Print "Form_Current 3:" + format$(Now, "hh:mm:ss")
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 syscmd 4, "Formularvorbereitung 3 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 Call KRAdd(frm, neurp, "Polyneuropathie", "G63.2", gesi, "G63", "G62", , , "vTextB", 105) ' Weitere Medikation
 If labPos <> 0 Then pnpflNeu = -1
 Call KRAdd(frm, autnp, "Autonome Neuropathie", "G99.0", gesi, "G99", , , , "vTextB", 105)  ' Weitere Medikation
 If labPos <> 0 Then anpflNeu = True
 Call KRAdd(frm, khk, "Koronare Herzerkrankung", "I25.9", gesi, "I25", "I21", , , "vTextB", 105)  ' Weitere Medikation
 Call KRAdd(frm, herzis, "Herzinsuffizienz", "I50.12", gesi, "I50", vNS, vNS, vNS, "vTextB", 105)  ' Weitere Medikation
 Call KRAdd(frm, avk, "AVK", "I73.9", gesi, "I74", "I70", "I73", , "vTextB", 105) ' Weitere Medikation
 If labPos <> 0 Then angflNeu = -1
 Call KRAdd(frm, urä, "Niereninsuffizienz", "N18.83", gesi, "N19", "N18", , , "vTextB", 105)  ' Weitere Medikation
 Call KRAdd(frm, antimyk, "Pilzinfektion", "B35.9", gesi, "B35.9", "B35", "B37", , "vTextB", 105) ' Weitere Medikation  ' B49, N77 fehlt noch: Vaginalmykose, B37.2 Onychomykose
 Call KRAdd(frm, hmg, "Hypercholesterinämie", "E78.0", gesi, "E78", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, fetts, "Fettstoffwechselstörung", "E78.9", gesi, "E78", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, hypt, "Arterielle Hypertonie", "I10.90", gesi, "I10", , , , "vTextB", 105)  ' Weitere Medikation
 Call KRAdd(frm, Hsre, "Hyperurikämie", "E79.0", gesi, "E79", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, glauk, "Glaukom", "H40.9", gesi, "H40", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, cold, "COPD", "J42", gesi, "J42", "J44", "M06.99", , "vTextB", 105) ' Weitere Medikation ' Cortison auch bei Rheuma
 Call KRAdd(frm, pros, "Prostatahyperplasie", "N40", gesi, "N40", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, hythy, "Hyperthyreose", "E05.9", gesi, "E05", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, stru, "Struma (nach Med.)", "E04.9", Va, "E04", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, ostp, "Osteoporose", "M81.99", gesi, "M81", "M82", , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, pani, "Pankreasinsuffizienz", "K86.8", gesi, "K86", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, park, "Parkinson", "G20", gesi, "G20", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, vari, "Varikose", "I83.9", gesi, "I83", "I86", , , "vTextB", 105) ' Weitere Medikation
 GebDat = IIf(IsNull(frm.anaRS!GebDat), 0, frm.anaRS!GebDat)
' Alter = (Date - GebDat) * 2.73792574745373E-03 ' 1/365,24
 Alter = AlterBei(Date, GebDat)
 If Alter > 45 Then Call KRAdd(frm, östr, "Wechselbeschwerden", "N95.9", gesi, "N95", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, antidep, "Depression", "F32.9", gesi, "F32", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, antidem, "Demenz", "F03", gesi, "F0", "G30", , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, antiep, "Epilepsie", "G40.9", gesi, "G40", "G41", , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, antipern, "Vitamin-B12-Mangel", "E53.8", gesi, "G63.4", "D51", "E53.8", , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, antiherp, "Herpes-Infektion", "B00.9", gesi, "B00", , , , "vTextB", 105) ' Weitere Medikation
 Call KRAdd(frm, Antikoag, "Antikoagulantien-Therapie", "Z92.1", gesi, , , , "vTextB", 105) ' Weitere Medikation
 
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 syscmd 4, "Formularvorbereitung 4 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
' Anamnesebogen
 Dim NPGru$
 NPGru = LegNPFest(Pat_ID, frm)
 
 If NPGru Like "->*" Then Call KRAdd(frm, Mid$(NPGru, 3), "Polyneuropathie", "G63.2", gesi, "G62", "G63")
 If labPos <> 0 Then pnpflNeu = -1
'  IF NPGrund = "KW" OR NPGrund = "MER" THEN
 If NPGru Like "V.*" Then Call KRAdd(frm, Mid$(NPGru, 3), "Polyneuropathie", "G63.2", Va, "G62", "G63")
 Call LegNPFest(Pat_ID, frm)
 
 Dim verform$, BewEinsch$, WSwo%
 verform = IIf(IsNull(frm.anaRS!Verformungen), vNS, LCase$(frm.anaRS!Verformungen))
 BewEinsch = IIf(IsNull(frm.anaRS!Bewegungseinschränkungen), vNS, LCase$(frm.anaRS!Bewegungseinschränkungen))
 Dim obBewEinsch%, obverform%
 obverform = 0
 obBewEinsch = 0
 If Left$(verform, 1) = "j" Then
  obverform = -1
  Call doGilbE(frm, "M14", "vTextB", 92) ' Bewegungseinschränkungen
 End If
 If obNein(BewEinsch) = 0 Then
  obBewEinsch = -1
  Call doGilbE(frm, "M14", "vTextB", 99)
 End If
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 If obverform Or obBewEinsch Then '(BewEinsch <> "" AND BewEinsch <> "n" AND BewEinsch <> "-" AND instrb(BewEinsch, "WS-Syndrom") = 0 AND (Not BewEinsch LIKE "o.B?") AND (Not BewEinsch LIKE "o. B?")) THEN
  Call KRAdd(frm, "Fragebogen", "Diab. Arthropathie", "M14.2-", Va, "M14") ' Bewegungseinschränkungen
  If InStrB(BewEinsch, "Knie") <> 0 Or InStrB(BewEinsch, "Gonarth") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Kniegelenksarthrose", "M17.9", Va, "M17", , , , "vTextB", 99) ' Bewegungseinschränkungen
  End If
  If InStrB(BewEinsch, "Hüft") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Hüftgelenksarthrose", "M16.9", gesi, "M16", , , , "vTextB", 99) ' Bewegungseinschränkungen
  End If
  If InStrB(BewEinsch, "Schulter") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Schultergelenksarthrose", "M19.21", Va, "M19", , , , "vTextB", 99) ' Bewegungseinschränkungen
  End If
  WSwo = 0
  If InStrB(BewEinsch, "lws") <> 0 Or InStrB(BewEinsch, "Lend") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "LWS-Syndrom", "M47.96", Va, "M47.96", "M54.16", "M47.8", "M53", "vTextB", 99) ' Bewegungseinschränkungen
    WSwo = 1
  End If
  If InStrB(BewEinsch, "bws") <> 0 Or InStrB(BewEinsch, "Brustwirbel") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "BWS-Syndrom", "M47.94", Va, "M47.94", "M54.14", "M47.8", "M53", "vTextB", 99) ' Bewegungseinschränkungen
    WSwo = 2
  End If
  If InStrB(BewEinsch, "hws") <> 0 Or InStrB(BewEinsch, "Halswirbel") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "HWS-Syndrom", "M47.92", Va, "M47.92", "M54.2", "M47.8", "M53", "vTextB", 99) ' Bewegungseinschränkungen
    WSwo = 3
  End If
  If InStrB(BewEinsch, "WS") <> 0 And WSwo = 0 Then
    Call KRAdd(frm, "Fragebogen", "Wirbelsäulen-Syndrom", "M47.99", Va, "M47", "M54", "M53", , "vTextB", 99) ' Bewegungseinschränkungen
    WSwo = 3
  End If
  If InStrB(BewEinsch, "Bandsch") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Diskopathie", "M51.8", Va, "M51", , , , "vTextB", 99) ' Bewegungseinschränkungen
  End If
 End If
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 Dim hkGr$
 hkGr = hkGrund(Pat_ID)
 If hkGr <> vNS Then
  Call KRAdd(frm, hkGr, "Hyperkeratose", "L84", gesi, "L85.9", "L84", , , "vTextB", IIf(hkGr = "Frgbg.", 91, 109)) ' Druckstellen / Hyperkeratosen
 End If ' LEFT(druckst, 1) = "j" THEN
 
 Dim EntlSt$, autGrund$, obAutNP, HarnB
 obAutNP = 0
 HarnB = 0
 EntlSt = IIf(IsNull(frm.anaRS![Entleerungsstörungen Magen]), vNS, LCase$(frm.anaRS![Entleerungsstörungen Magen]))
 If Left$(EntlSt, 1) = "j" Then
  obAutNP = obAutNP - 1
  Call doGilbE(frm, "G99", "vTextB", 95) ' Entleerungsstörungen Magen
 End If
 EntlSt = IIf(IsNull(frm.anaRS![Entleerungsstörungen Harnblase]), vNS, LCase$(frm.anaRS![Entleerungsstörungen Harnblase]))
 If Left$(EntlSt, 1) = "j" Then
  obAutNP = obAutNP - 1
  Call doGilbE(frm, "G99", "vTextB", 96) ' Entleerungsstörungen Harnblase
  HarnB = -1
 End If
 EntlSt = IIf(IsNull(frm.anaRS!Sexualstörung), vNS, LCase$(frm.anaRS!Sexualstörung))
 If Left$(EntlSt, 1) = "j" Then
  Call doGilbE(frm, "G99", "vTextB", 100) ' Sexualstörung
  obAutNP = obAutNP - 1
 End If
 If obAutNP <> 0 And Not (obAutNP And HarnB) Then ' Nur Harnblase lieber nicht
  Call KRAdd(frm, "Anamn", "Autonome Neuropathie", "G99.0", IIf(obAutNP = -1, Va, gesi), "G99")
  If labPos Then anpflNeu = True
 End If
 
 oblhNeu = obLH(Pat_ID, frm)
 
 If oblhNeu Then
  Call KRAdd(frm, "Befund", "Liphypertrophien", "L94.8", gesi, "L94", "E88")
 End If
 
 Dim ulcer$
 If IsNull(frm.anaRS!Ulcera) Then
  ulcer = vNS
 Else
  ulcer = frm.anaRS!Ulcera
 End If
 ulcer = Trim$(ulcer)
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 If Not obNein(ulcer) Then
  Call KRAdd(frm, "Unters", "Diabetisches Ulcus", "L89.18", gesi, "L89", , , , "vTextB", 110) ' "Ulcera"
  If labPos <> 0 Then dfsflNeu = True
 Else
  Dim Geschwür$
  Geschwür = IIf(IsNull(frm.anaRS!Geschwür), vNS, LCase$(frm.anaRS!Geschwür))
  If Geschwür <> vNS And Geschwür <> "-" And Geschwür <> "n" Then
   Call KRAdd(frm, "Anamn", "Diabetisches Ulcus", "L89.18", gesi, "L89", vNS, vNS, vNS, "vTextB", 86) ' "Geschwür"
   If labPos <> 0 Then dfsflNeu = True
  End If
 End If

 
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 syscmd 4, "Formularvorbereitung 5 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 Dim PStatPath% ' Pulsstatus pathologisch
 PStatPath = 0
 PStatPath = PStatus(Pat_ID, frm)
 
 If PStatPath Then
  Call KRAdd(frm, "Pulse", "Periphere arterielle Verschlußkrankheit", "I70.20", gesi, "I73", "I74", "I70", , "vTextB", 123) ' Puls Atp
  If labPos <> 0 Then angflNeu = -1
 End If
 
 Dim HSchw$, Lunge$, herzin%, HIGrund$
 HSchw = IIf(IsNull(frm.anaRS!Herzschwäche), vNS, LCase$(frm.anaRS!Herzschwäche))
 If InStrB(HSchw, "j") <> 0 Or InStrB(HSchw, "dyspn") <> 0 Then
  herzin = True
  HIGrund = "Frgb"
 End If
 Lunge = IIf(IsNull(frm.anaRS!Lunge), vNS, LCase$(frm.anaRS!Lunge))
 If InStrB(Lunge, "feucht") <> 0 Then
  herzin = True
  HIGrund = IIf(LenB(HIGrund) = 0, vNS, HIGrund + ", ") + "Ausk"
 End If
 If herzin Then
  Call KRAdd(frm, HIGrund, "Herzinsuffizienz", "I50.12", Va, "I50", vNS, vNS, vNS, "vTextB", IIf(HIGrund = "Frgb", 80, 127)) ' Herzschwäche / Lunge
 End If
 
 Dim WS$, obWS%, obBewEin%
 WS = IIf(IsNull(frm.anaRS!WS), vNS, LCase$(frm.anaRS!WS))
 If (InStrB(WS, "ks") <> 0 And InStrB(WS, "kein") = 0) Then
  Call doGilbE(frm, "M53.99", "vTextB", 129) ' WS
  obWS = -1
 End If
 If obWS <> 0 Or obBewEin Then
  Call KRAdd(frm, IIf(obWS, "Unters", "Anamn"), "Wirbelsäulensyndrom", "M53.99", gesi, "M53", "M54", "M47")
 End If
 
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 Dim Strum$
 If Not IsNull(frm.anaRS!SD) Then
  Strum = LCase$(frm.anaRS!SD)
 End If
 If InStrB(Strum, "stru") <> 0 Then
  If InStrB(Strum, "v.a") <> 0 Then
   Call KRAdd(frm, "Unters", "Struma", "E04.9", Va, "E04", , , , "vTextB", 131) ' SD
  Else
   Call KRAdd(frm, "Unters", "Struma", "E04.9", gesi, "E04", , , , "vTextB", 131) ' SD
  End If
 End If
 
 Dim weit$, lcWA$, lcAB$, lcSU$
 If IsNull(frm.anaRS![weitere anamnese]) Then
  lcWA = vNS
 Else
  lcWA = LCase$(frm.anaRS![weitere anamnese])
 End If
 If IsNull(frm.anaRS![augensp befund]) Then
  lcAB = vNS
 Else
  lcAB = LCase$(frm.anaRS![augensp befund])
 End If
 If IsNull(frm.anaRS![sehminderung unbehebbar]) Then
  lcSU = vNS
 Else
  lcSU = LCase$(frm.anaRS![sehminderung unbehebbar])
 End If
 If InStrB(lcWA, "psoriasis") <> 0 And InStrB(LCase$(frm.vTextB(147)), "psoriasis") = 0 Then
  Call KRAdd(frm, "Anamn", "Psoriasis", "L40.9", gesi, "L40", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 If InStrB(lcWA, "osteopor") <> 0 And InStrB(LCase$(frm.vTextB(147)), "m81") = 0 Then
  Call KRAdd(frm, "Anamn", "Osteoporose", "M81.99", gesi, "M81", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(lcWA, "spinalst") <> 0 And InStrB(LCase$(frm.vTextB(147)), "m48") = 0 Then
  Call KRAdd(frm, "Anamn", "Spinalstenose", "M48.09", gesi, "M48", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If (InStrB(lcWA, "neurodermitis") <> 0 Or InStrB(LCase$(frm.anaRS![Folgeerkrankungen Haut]), "neurodermitis") <> 0) And InStrB(LCase$(frm.vTextB(147)), "l20.8") = 0 Then
  Call KRAdd(frm, "Anamn", "Neurodermitis", "L20.8", gesi, "L20.8", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If (InStrB(lcWA, "rhythmusstör") <> 0 Or InStrB(lcWA, "sm-imp") <> 0 Or InStrB(LCase$(frm.anaRS![Herzkrankheit Beschreibung]), "rhythmusstö") <> 0) And InStrB(LCase$(frm.vTextB(147)), "i49") = 0 And InStrB(LCase$(frm.vTextB(147)), "i48") = 0 Then
  Call KRAdd(frm, "Anamn", "Herzrhythmusstörungen", "I49.9", gesi, "I49", "I48", , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If (InStrB(lcWA, "vorhoffli") <> 0 Or InStrB(LCase$(frm.anaRS![Herzkrankheit Beschreibung]), "vorhoffli") <> 0) And InStrB(LCase$(frm.vTextB(147)), "i48") = 0 Then
  Call KRAdd(frm, "Anamn", "Vorhofflimmern", "I48.2", gesi, "I48", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(LCase$(frm.anaRS![Entleerungsstörungen Harnblase]), "j") <> 0 And InStrB(frm.vTextB(147), "N31") = 0 And InStrB(frm.vTextB(147), "N39") = 0 Then
  Call KRAdd(frm, "Anamn", "Störung der Harnblase", "N31.9", gesi, "N31", "N39", , , IIf(InStrB(LCase$(frm.anaRS![Entleerungsstörungen Harnblase]), "j"), "Entleerungsstörungen Harnblase", vNS))
 End If
 If (InStrB(lcWA, "cold") <> 0 Or InStrB(lcWA, "cole") <> 0 Or InStrB(lcWA, "copd") <> 0) And InStrB(LCase$(frm.vTextB(147)), "j42") = 0 Then
  Call KRAdd(frm, "Anamn", "Chronische obstruktive Atemwegserkrankung", "J42", gesi, "J42", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(lcWA, "asthma") <> 0 And InStrB(LCase$(frm.vTextB(147)), "j45") = 0 Then
  Call KRAdd(frm, "Anamn", "Asthma bronchiale", "J45.9", gesi, "J45", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(LCase$(frm.anaRS!Herz), "ystolikum") <> 0 And InStrB(LCase$(frm.vTextB(147)), "R01") = 0 And InStrB(LCase$(frm.vTextB(147)), "I34") = 0 And InStrB(LCase$(frm.vTextB(147)), "I35") = 0 Then
  Call KRAdd(frm, "Befd", "Systolikum", "R01.1", gesi, "R01.1", "I34", "I35", , "Herz")
 End If
 If InStrB(lcWA, "allergie") <> 0 And InStrB(LCase$(frm.vTextB(147)), "t78") = 0 Then
  Call KRAdd(frm, "Anamn", "Allergie", "T78.4", gesi, "T78", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If (InStrB(lcWA, "cts") <> 0 Or InStrB(lcWA, "karpaltunnel") <> 0) And InStrB(LCase$(frm.vTextB(147)), "g56.0") = 0 Then
  Call KRAdd(frm, "Anamn", "Karpaltunnelsyndrom", "G56.0", gesi, "G56.0", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(lcWA, "kreuzschmerzen") <> 0 And InStrB(LCase$(frm.vTextB(147)), "m54") = 0 Then
  Call KRAdd(frm, "Anamn", "LWS-Syndrom", "M54.16", gesi, "M54", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(lcWA, "byp") <> 0 And InStrB(LCase$(frm.vTextB(147)), "z95") = 0 Then
  Call KRAdd(frm, "Anamn", "Bypaß peripher", "Z95.88", gesi, "Z95.88", , , , "vTextB", 102) ' "Weitere Anamnese")
  Call KRAdd(frm, "Anamn", "Bypaß kardial", "Z95.1", gesi, "Z95.1", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If InStrB(lcWA, "acvb") <> 0 And Not frm.anaRS![Bypaß kardial] And InStrB(LCase$(frm.vTextB(147)), "z95.1") = 0 Then
  Call KRAdd(frm, "Anamn", "Bypaß kardial", "Z95.1", gesi, "Z95.1", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
 If (InStrB(lcWA, "glaukom") <> 0 Or InStrB(LCase$(frm.anaRS![sehminderung unbehebbar]), "glaukom") <> 0 Or InStrB(lcAB, "glaukom") <> 0 Or _
    InStrB(lcWA, "grüner star") <> 0 Or InStrB(LCase$(frm.anaRS![sehminderung unbehebbar]), "grüner star") <> 0 Or InStrB(lcAB, "grüner star") <> 0) And InStrB(LCase$(frm.vTextB(147)), "glaukom") = 0 Then
  Call KRAdd(frm, "Anamn", "Glaukom", "H40.9", gesi, "H40", , , , "vTextB", 102) ' "Weitere Anamnese")
 End If
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 Dim catpos%, catzn%, obcatzn%
 catpos = 0
 catzn = 0
 obcatzn = 0
 Dim AugFeld$, AugNr%
 For i = 1 To 4
  Select Case i
   Case 1: AugFeld = "vTextB": AugNr = 102 ' AugFeld = "Weitere Anamnese"
   Case 2: AugFeld = "vTextB": AugNr = 68 '  "Sehminderung unbehebbar"
   Case 3: AugFeld = "vTextB": AugNr = 66 '"Augensp Befund"
   Case 4: AugFeld = "vTextB": AugNr = 67 '"Netzhaut gelasert"
  End Select
  If catpos = 0 And Not IsNull(frm(AugFeld)(AugNr)) Then
   catpos = InStr(frm(AugFeld)(AugNr).DataField, "atarac")
   If catpos = 0 Then catpos = InStr(frm(AugFeld)(AugNr).DataField, "atarak")
   If catpos = 0 Then catpos = InStr(frm(AugFeld)(AugNr).DataField, "grauer star")
   If catpos > 0 Then
    Call doGilbE(frm, "H26.9", AugFeld, AugNr)
    catzn = InStr(lcAB, "z.n.")
    If catzn > 0 And catzn < catpos Then
     obcatzn = -1
    Else
     catzn = InStr(frm(AugFeld), "op")
     If catzn > 0 And catzn < catpos + 15 Then
      obcatzn = -1
     End If
    End If
   End If
  End If
 Next i
 If catpos > 0 Then
  Call KRAdd(frm, "Anamn", "Diab. Katarakt", "H28.0", IIf(catzn > 0, Zn, gesi), "H26", "H28")
 End If
 
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 Dim zähn$
 zähn = IIf(IsNull(frm.anaRS!Zähne), vNS, LCase$(frm.anaRS!Zähne))
 If InStrB(zähn, "ungsbed") <> 0 Then
  Call KRAdd(frm, "Unters", "Kariöse Zähne", "K02.9", gesi, "K02", , , , "vTextB", 133) ' Zähne
 End If
 
 Dim beinöd$
 If IsNull(frm.anaRS!BeinödVen) Then
  beinöd = vNS
 Else
  beinöd = LCase$(frm.anaRS!BeinödVen)
 End If
 If (InStrB(beinöd, "ödem") <> 0 And Not beinöd Like "*kein ödem*" And Not beinöd Like "*keine ödem*") Or InStrB(beinöd, "gering") <> 0 Or InStrB(beinöd, "deutl") <> 0 Or InStrB(beinöd, "interm") <> 0 Or InStrB(beinöd, "stauungsderm") <> 0 Then
  Call KRAdd(frm, "Unters", "Beinödeme", "R60.0", gesi, "R60", , , , "vTextB", 161) ' BeinödVen
 End If
' beinöd = IIF(ISNULL(frm.anaRS!Beinbefund), vns, LCase$(frm.anaRS!Beinbefund))
 If IsNull(frm.anaRS!Beinbefund) Then
  beinöd = vNS
 Else
  beinöd = LCase$(frm.anaRS!Beinbefund)
 End If
 If (InStrB(beinöd, "ödem") <> 0 And Not beinöd Like "*kein ödem*" And Not beinöd Like "*keine ödem*") Or InStrB(beinöd, "gering") <> 0 Or InStrB(beinöd, "stauungsderm") <> 0 Then
  Call KRAdd(frm, "Unters", "Beinödeme", "R60.0", gesi, "R60", , , , "vTextB", 141) ' BeinBefund
 End If
 
 Dim Beinbefund$
 If Not IsNull(frm.anaRS!Beinbefund) Then
  Beinbefund = LCase$(frm.anaRS!Beinbefund)
 End If
 If Not IsNull(frm.anaRS!Hyperkeratosen) Then
  Beinbefund = Beinbefund & " " & LCase$(frm.anaRS!Hyperkeratosen)
 End If
 If Not IsNull(frm.anaRS!Ulcera) Then
  Beinbefund = Beinbefund & " " & LCase$(frm.anaRS!Ulcera)
 End If
 
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 If InStrB(Beinbefund, "hallux valgus") <> 0 Then
  Call KRAdd(frm, "Beinbefund", "Hallux valgus", "M20.1", gesi, "M20.1", , , , "vTextB", 141)
 End If

 If InStrB(Beinbefund, "myk") <> 0 Or InStrB(Beinbefund, "tinea") <> 0 Then
  If InStrB(Beinbefund, "nagel") <> 0 Then
   Call KRAdd(frm, "Unters", "Onychomykose", "B37.2", Va, "B37", , , , "vTextB", 141) ' Beinbefund
  Else
   Call KRAdd(frm, "Unters", "Fußmykose", "B35.3", Va, "B35", "B37", , , "vTextB", 141) ' Beinbefund
  End If
 End If
 If InStrB(Beinbefund, "necrob") <> 0 Or InStrB(Beinbefund, "nekrob") <> 0 Then
  Call KRAdd(frm, "Unters", "Necrobiosis lipoidica", "L92.1", gesi, "L92.1", "L99.8", , , "vTextB", 141) ' Beinbefund
 End If
 If (InStrB(Beinbefund, "psoriasis") <> 0 Or InStrB(LCase$(frm.anaRS![Folgeerkrankungen Haut]), "psoriasis") <> 0) And InStrB(LCase$(frm.vTextB(147)), "psoriasis") = 0 Then
  Call KRAdd(frm, "Unters", "Psoriasis", "L40.9", gesi, "L40", , , , "vTextB", 141) ' Beinbefund
 End If
 If InStrB(Beinbefund, "senk-") <> 0 Or InStrB(Beinbefund, "senkfu") <> 0 Or InStrB(Beinbefund, "platt") <> 0 Or InStrB(Beinbefund, "planus") <> 0 Then
  Call KRAdd(frm, "Unters", "Senk-Spreizfuß", "M21.4", gesi, "M21.4", "Q66.8", , , "vTextB", 141) ' Beinbefund
 End If
 If InStrB(Beinbefund, "krallenzeh") <> 0 Then
  Call KRAdd(frm, "Unters", "Krallenzehen", "Q66.8", gesi, "Q66.8", , , , "vTextB", 141) ' Beinbefund
 End If
 If InStrB(Beinbefund, "hohlfu") <> 0 Then
  Call KRAdd(frm, "Unters", "Hohlfuß", "Q66.7", gesi, "Q66.7", , , , "vTextB", 141) ' Beinbefund
 End If
 
 
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If

 syscmd 4, "Formularvorbereitung 6 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
' IF 1 = 0 THEN
'  frm.MDi = frm.MDi & vbCrLf
'  MDIn = MDIn + 1
' END IF
  
 Alter = 0
 Pat_ID = Pat_ID
 obWeib = IIf(frm.anaRS!anrede = "Herr", False, True)
 'Set rlau = TabÖff("Labor", "WertSuch")
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", vor do_Form_Current2)"
#End If
 Call do_Form_Current2(frm, Hsre, lcAB, PStatPath, pnpflNeu, KZahlNeu, anpflNeu, angflNeu, bgwflNeu, retflNeu, dfsflNeu, oblhNeu)
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", nach do_Form_Current2)"
#End If
End If
Call KrSchluß(frm)
Dim Ther1$
syscmd 4, "Formularvorbereitung 15 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
'Ther1 = TherUmw(TherArt(Pat_id, -1))
'If Ther1 <> frm.anaRS!Ther1 THEN frm.vTextB(166) = Ther1
 frm.Erklärung = vNS
 frm.Inhalt = vNS
'Debug.Print "Form_Current 5:" + format$(Now, "hh:mm:ss")
#If mitab Then
On Error Resume Next
Call ddsono(frm)
On Error GoTo fehler
Dim rs As New ADODB.Recordset
myFrag rs, "SELECT quartal, auftrag, verdacht, befund FROM `faelle` WHERE pat_id = " & Pat_ID & " AND ((NOT ISNULL(auftrag) AND auftrag <> '') OR (NOT ISNULL(verdacht) AND verdacht <> '') OR (NOT ISNULL(befund) AND befund <> '')) ORDER BY bhfb DESC"
If Not rs.BOF Then
 frm.Auftrag = rs!Auftrag
 frm.Verdacht = rs!Verdacht
 frm.Befund = rs!Befund
Else
 frm.Auftrag = vNS
 frm.Verdacht = vNS
 frm.Befund = vNS
End If
#End If
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", vor DBCnS)"
#End If
frm.vTextB(182) = DBCnS ' DBCn.ConnectionString
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", nach DBCnS)"
#End If
syscmd 5
Tüt 1740, 20
TütSoundkarte
kurzvorEnde:
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
 Close #313
#End If
frm.obStumm = False
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Form_Current_AnBog/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Form_Current_AnBog

Function do_Form_Current2(frm As AnBog, Hsre$, lcAB$, PStatPath%, pnpflNeu%, KZahlNeu%, anpflNeu%, angflNeu%, bgwflNeu%, retflNeu%, dfsflNeu%, oblhNeu%)
 Dim Aldo_p!
' Dim Krea!, Kre02!, Krea02!, Kres!, Creat!, Hst!
 Dim Adrenu!
' Dim AlbCre!, AlbKre!, AlbQ!, Album!, Albup!
'Dim B12!
 Dim Arq!, Amyl!, Lipase!, ANA!, Andro!, ap!, APCI_CP!, ATIII_K!, _
     BiliG!, Borg!, Ca!, Ca199_S!, cAnca!, Carigg!, Carigm!, Cea_k!, CHE!, Chltra!, Chltrg!, Chol!, Cicc1d!, _
     Cicc1q!, CKmb!, CKNAC!, CMVIGM!, COERUL!, CORS!, CPEP!, CREACL!, CRP!, CUS!, DIGIT!, DIGOX!, EBVM!, FERR!, FT3_K!, FT4_K!, GAGT!, GGT!, GPT!, GLU!, HAPTO!, HAVm_K!, Hb!, Hk!, HbA1c!, HBC!, HBSAG_K!, HCV!, HCYP!, HGEM_s!, HS!, MG_K!, IA2AK!, IAK_s!, Na!, IGE_K!, IGF1_S!, INSELA!, INSELQ!, k!, LACT_E!, LDL!, LDLB!, LDLH!, LDLH01!, LDLC!, LDLLG!, LDLS!, Leuko!, LPa!, MAK!, MCH!, NH4!, P_ANCQ_S!, PANC_s!, PANELF!, PCAK_cp!, PHOS!, PO4!, PROL_K!, pSa!, PSAK_cp!, rf!, RFQ!, ST24!, THRM!, THROMB!, TRAK_k!, TSH!, TSHL_k!, TSBF!, VITD25_s!, WACHSH_s!, CPEP_S!
 Dim ccp%
' Diagnosen
 Dim nieflNeu%, nasflNeu%
 On Error GoTo fehler
 nieflNeu = 0
 nasflNeu = 0

 syscmd 4, "Formularvorbereitung 6a " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 'sql = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" AS NB FROM (" & laborAbfr & " WHERE pat_id = " & Pat_id & ") AS labor UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar, NB " + _
 "FROM `laborxus` LEFT JOIN `laborxwert` ON `laborxus`.RefNr=`laborxwert`.RefNr " + _
 "WHERE pat_id = " + CStr(Pat_id) + " AND NOT EXISTS (SELECT * FROM `laborneu` WHERE pat_id = " & CStr(Pat_id) & " AND abkü = laborxwert.Abkü AND wert = laborxwert.wert AND zeitpunkt > laborxus.Eingang-3 AND zeitpunkt <  laborxus.Eingang+6) "
 sql = "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_ID & " UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_ID & ") i GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb"
 sql1 = " ORDER BY zeitpunkt DESC"
 
' SET raLau = Dtb.OpenRecordset(sql)
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", vor LabWert)"
#End If
 Set raLau = Nothing
' raLau.Open LCase$(sql) & sql1, DBCn, adOpenDynamic, adLockReadOnly
 myFrag raLau, LCase$(sql) & sql1, adOpenDynamic
'Set rLaU = Dtb.OpenRecordset("SELECT * FROM `" + QMdbAkt + "`.LaborUNION WHERE pat_id = " + CStr(pat_id) + " ORDER BY zeitpunkt DESC")
 syscmd 4, "Formularvorbereitung 6a1 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 ' 'CREAT','KRE02','KREA','KREA02','KRES'
#If True Then
 Call LabwertA(frm, LA_Krea, ">", 1.44, "Nephropathie", "N08.3", Va, "mg/dl", "N08.3", "N08.3", 1.8, 1.1)
#Else
' Krea02 = LabWert(frm, "Krea02", ">", 1.44, "Nephropathie", "N08.3", Va, "mg/dl", "N08.3", "N08.3", 1.8, 1.1)
' IF labPos <> 0 THEN nieflNeu = -1
' Creat = LabWert(frm, "CREAT", ">", 1.44, "Nephropathie", "N08.3", Va, "mg/dl", "N08.3", "N08.3", 1.8, 1.1)
' IF labPos <> 0 THEN nieflNeu = -1
' Kre02 = LabWert(frm, "Kre02", ">", 1.44, "Nephropathie", "N08.3", Va, "mg/dl", "N08.3", "N08.3", 1.8, 1.1)
' IF labPos <> 0 THEN nieflNeu = -1
' Krea = LabWert(frm, "Krea", ">", 1.44, "Nephropathie", "N08.3", Va, "mg/dl", "N08.3", "N08.3", 1.8, 1.1)
' IF labPos <> 0 THEN nieflNeu = -1
' Kres = LabWert(frm, "Kres", ">", 1.44, "Nephropathie", "N08.3", Va, "mg/dl", "N08.3", "N08.3", 1.8, 1.1)
#End If
 If labPos <> 0 Then nieflNeu = -1
#If True Then
 Call LabwertA(frm, LA_Krea, ">", 1.44, "Niereninsuffizienz", "N19", gesi, "mg/dl", "N18", "N19", 1.8, 1.35, "N17")
#Else
' Krea = LabWert(frm, "Krea", ">", 1.44, "Niereninsuffizienz", "N19", gesi, "mg/dl", "N18", "N19", 1.8, 1.35, "N17")
' Krea02 = LabWert(frm, "Krea02", ">", 1.44, "Niereninsuffizienz", "N19", gesi, "mg/dl", "N18", "N19", 1.8, 1.35, "N17")
' Kre02 = LabWert(frm, "Kre02", ">", 1.44, "Niereninsuffizienz", "N19", gesi, "mg/dl", "N18", "N19", 1.8, 1.35, "N17")
' Kres = LabWert(frm, "Kres", ">", 1.44, "Niereninsuffizienz", "N19", gesi, "mg/dl", "N18", "N19", 1.8, 1.35, "N17")
' Creat = LabWert(frm, "Creat", ">", 1.44, "Niereninsuffizienz", "N19", gesi, "mg/dl", "N18", "N19", 1.8, 1.35, "N17")
' Hst = LabWert(frm, "hst")
#End If
 Adrenu = LabWert(frm, "adrenu") '20 µg/24h
' 'ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP'
#If True Then
 Call LabwertA(frm, LA_AlbCre, ">", 30, "Nephropathie", "N08.3", gesi, "mg/gCrea", "N08.3", "N08.3", 50)
#Else
' AlbCre = LabWert(frm, "albcre", ">", 20, "Nephropathie", "N08.3", gesi, "mg/gCrea", "N08.3", "N08.3", 50)
' AlbKre = LabWert(frm, "albkre", ">", 20, "Nephropathie", "N08.3", gesi, "mg/g Krea", "N08.3", "N08.3", 50)
' AlbQ = LabWert(frm, "albq", ">", 20, "Nephropathie", "N08.3", gesi, "mg/g Krea", "N08.3", "N08.3", 50)
'' Album = LabWert(frm, "album", ">", 20, "Nephropathie", "N08.3", gesi, "mg/g Krea", "N08.3", "N08.3", 50) ' Einheit fließt nicht ein
' Albup = LabWert(frm, "albup", ">", 20, "Nephropathie", "N08.3", gesi, "mg/g Krea", "N08.3", "N08.3", 50)
#End If
 If labPos <> 0 Then nieflNeu = -1
 syscmd 4, "Formularvorbereitung 6b " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
' obkNeph = obKeineNephropathie(pat_id, obMakroAlb)
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", Mitte Labwert)"
#End If
 Dim UZahl%
 UZahl = Urineintraege(Pat_ID)
 If UZahl > 1 Then 'And Not obkNeph THEN
   Call KRAdd(frm, CStr(UZahl) + " Stix-eintraege", "Nephropathie", "N08.3", Va, "N08.3")
   If labPos <> 0 Then nieflNeu = -1
 End If
' Album = LabWert0(frm, "Album", "<", 35, "Hypalbuminämie", "E46", gesi, "g/l", "E46", "N04")
 Aldo_p = LabWert(frm, "aldo_p") ' 15
 Arq = LabWert(frm, "arq") '3,0
 Amyl = LabWert(frm, "amyl", ">", 100, "Pankreatitis", "K85.0", Va, "U/l", "K85")
 Lipase = LabWert(frm, "lipase", ">", 60, "Pankreatitis", "K85.0", Va, "U/l", "K85")
 ANA = LabWert(frm, "ana", ">", 80, "Lupus erythematodes o.ä.", "L93.0", Va, vNS, "L93", "H01")
 Andro = LabWert(frm, "andro") 'ca. 0,6-3,2/1,1-2,8 s. Internet
 ap = LabWert(frm, "ap") '104/129
 APCI_CP = LabWert0(frm, "APCI_CP", "<", 2.1, "APC-Resistenz", "D68.8", gesi, vNS, "D68", "I82.9", "I82.9")
 ATIII_K = LabWert0(frm, "AT III_K", "<", 75, "Antithrombinmangel", "D68.8", gesi, "%", "D68", "I82.9", 40)
#If True Then
 Call LabwertA(frm, LA_B12, "<", 250, "Vit.B12-Mangel", "G63.4", gesi, "pg/ml", "D51", "G63.4", 211, 250, "E53.8") ' "G63.4", "D51", "E53.8"
#Else
' B12 = LabWert0(frm, "b12", "<", 250, "Vit.B12-Mangel", "G63.4", gesi, "pg/ml", "D51", "G63.4", 211, 250, "E53.8") ' "G63.4", "D51", "E53.8"
#End If
 BiliG = LabWert(frm, "bilig", ">", 1.1, "Cholestase", "K83.1", Va, "mg/dl", "K83", "K74", 1.1, 1.1, "K71")
 If labPos <> 0 Then nasflNeu = -1
 Borg = LabWert(frm, "borg", ">", 6, "Borreliose", "A69.2", Va, "U/ml", "A68", "A69.2")
 Ca = LabWert0(frm, "ca", "<", 2.1, "Hypocalcämie", "E83.5", Va, "mmol/l", "E83.5", "E83.5", 2)
 Ca = LabWert(frm, "ca", ">", 2.6, "Hypercalcämie", "E83.5", gesi, "mmol/l", "E83.5", "E83.5", 2.7)
 Ca199_S = LabWert(frm, "CA199_S", ">", 38, "Pankreas-Carcinom", "C25.-", Va, "U/ml", "C25")
 cAnca = LabWert(frm, "C-ANCA", ">", 2, "Wegener/mPAN", "M31.3", Va, vNS, "M31.3", "M30") ' 2
 Carigg = LabWert(frm, "CARIGM", ">", 12, "Cardiolipin-Ak (IgG)", "D68.8", gesi, "U/ml", "A5", "I82.9", 12, 12, "D68.8")
 Carigm = LabWert(frm, "CARIGM", ">", 12, "Cardiolipin-Ak (IgG)", "D68.8", gesi, "U/ml", "A5", "I82.9", 12, 12, "D68.8")
 Cea_k = LabWert(frm, "CEA_K", ">", 5, "Carcinom", "C80", Va, "ng/ml", "C80", "C80", 10)
 CHE = LabWert0(frm, "CHE", "<", 5.3, "Lebersynthesestörung", "K74.6", Va, "kU/l", "K74", "K71", 5.3, 5.3, "K70")
 If labPos <> 0 Then nasflNeu = -1
 Chltra = LabWert(frm, "CHLTRA", ">", 1, "Trachom", "A71")
 Chltrg = LabWert(frm, "CHLTRG", ">", 1, "Trachom", "A71")
 Chol = LabWert(frm, "chol", ">", 200, "Hypercholesterinämie", "E78.0", gesi, "mg/dl", "E78", "E78")
 Cicc1d = LabWert(frm, "CIC-C1D", ">", 8, "Immunkomplexkranheit", "N05.8", Va, "µgHAG/ml")
 Cicc1q = LabWert(frm, "CIC-C1Q", ">", 10, "Immunkomplexkranheit", "N05.8", Va, "µgHAG/ml")
 CKmb = LabWert(frm, "CKMB", ">", 11.4, "Myokardinfarkt", "I25.2-", Va, "U/l", "I25", "I21", 10)
 CKNAC = LabWert(frm, "CKNAC", ">", 190, "Muskelerkrankung", "G72.9", Va, "U/l", "G72", 167)
 CMVIGM = LabWert(frm, "CMVIGM", ">", 0.9, "Cytomegalievirusinfektion", "B25", Va, vNS, "B25")
 COERUL = LabWert0(frm, "COERUL", "<", 0.2, "M.Wilson", "E83.0", Va, "g/l", "E83")
' rLa.Seek "=", Pat_id, "COR-S"
' raLau.FindFirst "abkü = ""COR-S"""
 Set raLau = Nothing
' raLau.Open sql & " AND abkü = ""COR_S"" " & sql1, DBCn, adOpenDynamic, adLockReadOnly
  myFrag raLau, sql & " AND abkü = ""COR_S"" " & sql1, adOpenDynamic
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", nach Labwert)"
#End If
 syscmd 4, "Formularvorbereitung 7 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
' Debug.Print "Form_Current 4:" + format$(Now, "hh:mm:ss")
 ' zwei Tage hintereinander, umgekehrte Zeitreihenfolge
 syscmd 4, "Formularvorbereitung 6c " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 If Not raLau.EOF Then
  If raLau!Wert < 4.3 Then
   raLau.Move 1
   If Not raLau.EOF() Then
    If raLau!Pat_ID = Pat_ID And raLau!Abkü = "COR-S" Then
     If raLau!Wert < 4.3 Then
      CORS = LabWert0(frm, "COR-S", "<", 4.3, "M.Addison", "E27.1", "µg/dl", "E27")
     End If
    End If
   End If
 End If
' ralau.Seek "=", Pat_id, "COR-S"
' raLau.FindFirst "abkü = ""COR-S"""
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 Set raLau = Nothing
' raLau.Open sql & " AND abkü = ""COR_S"" " & sql1, DBCn, adOpenDynamic, adLockReadOnly
 myFrag raLau, sql & " AND abkü = ""COR_S"" " & sql1, adOpenDynamic
 
 Dim Zp As Date, VorWert!
 VorWert = 0
 If Not raLau.EOF Then
  Do While Not raLau.EOF
   If raLau!Pat_ID <> Pat_ID Or raLau!Abkü <> "COR-S" Then Exit Do
    If raLau!Wert > 5 Then
     If VorWert > 20 And Zp = raLau!Zeitpunkt - 1 Then
      CORS = LabWert(frm, "COR-S", ">", 22.4, "M.Cushing", "E24.-", Va, "µg/dl", "E24")
     Else
      Zp = DateValue(raLau!Zeitpunkt)
      Dim VWStr$
      VWStr = REPLACE$(raLau!Wert, "<", vNS)
      If IsNumeric(VWStr) Then VorWert = VWStr Else VorWert = 0
     End If
    End If
   raLau.Move 1
  Loop
 End If
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 'CORT2 = labwert(frm,"CORT2", ">", 3, "M.Cushing", gesi,"µg/dl")
 'CORT2 = labwert(frm,"CORT2", "<", 20, "M.Addison", gesi,"µg/dl")
 CPEP_S = LabWert(frm, "C PEP_S", ">", 0.1, "Restsekretion", vNS, gesi, "ng/ml")
 CREACL = LabWert(frm, "CREACL")
 CRP = LabWert(frm, "CRP", ">", 5, "Entzündung", vNS, Va, "mg/l", "R70.0")
 CUS = LabWert0(frm, "CU-S", "<", 70, "M.Wilson", "E83.0", Va, "µg/dl", "E83.0", "E83.0", 85)
 DIGIT = LabWert(frm, "DIGIT", ">", 30, "Digitoxinintoxikation", "T46.0", Va, "ng/ml")
 DIGOX = LabWert(frm, "DIGOX", ">", 2, "Digoxinintoxikation", "T46.0", Va, "ng/ml")
 EBVM = LabWert(frm, "EBVM", ">", 1, "EBV-Virusinfektion", "B27.0", Va, vNS)
 FERR = LabWert(frm, "FERR", ">", 600, "Hämochromatose", "E83.1", Va, "µg/l", "E83.1")
 FT3_K = LabWert(frm, "FT3_K", ">", 4.2, "Hyperthyreose", "E05.9", gesi, "pg/ml", "E05")
 FT4_K = LabWert(frm, "FT4_K", ">", 1.7, "Hyperthyreose", "E05.9", gesi, "ng/dl", "E05")
 FT4_K = LabWert0(frm, "FT4_K", "<", 0.8, "Hypothyreose", "E03.9", gesi, "ng/dl", "E03", "E89", 0.8, 0.8, "E02")
 GAGT = LabWert(frm, "GAGT", ">", 66, "Hepatopathie", "K77.8", gesi, "U/l", "K71", "K76", 66, 39, "K77")
 If labPos <> 0 Then nasflNeu = -1
 GGT = LabWert(frm, "GGT", ">", 66, "Hepatopathie", "K77.8", gesi, "U/l", "K71", "K76", 66, 39, "K77")
 If labPos <> 0 Then nasflNeu = -1
 GPT = LabWert(frm, "GPT", ">", 50, "Hepatopathie", "K77.8", gesi, "U/l", "K71", "K76", 50, 35, "K77")
 If labPos <> 0 Then nasflNeu = -1
 GLU = LabWert(frm, "GLU", ">", 200, "Diabetes mellitus", "E11.9", gesi, "mg/dl", "E10", "E11", 130, 200, "E11")
 HAPTO = LabWert0(frm, "HAPTO", "<", 0.3, "Hämolyse", "D59.9", gesi, "g/l", "D5")
 HAVm_K = LabWert(frm, "HAV M_K", ">", 0.1, "Hepatitis A", "B15.-", gesi, vNS, "B15")
 Hb = LabWert0(frm, "HB", "<", 14, "Anämie", "K92.2", gesi, "g/dl", "D64", "D5", 12.6, 12, "K9")
 Hb = LabWert(frm, "HB", ">", 17.4, "Polyglobulie", "J44.9-", gesi, "g/dl", "D45", "J44.9-", 17.4, 16)
 Hk = LabWert(frm, "HK", ">", 54, "Polyglobulie", "J44.9-", gesi, vNS, "D45", "J44.9-", 54, 52)
' HbA1c = LabWert(frm, "HBA1C", ">=", 6.5, "schlechte Einstellung", vNS, gesi, "%", ".7" & IIF(InStrB(frm.vTextB(147), "L89") = 0, "3", "5"), ".61")
' HbA1c = LabWert0(frm, "HBA1C", "<", 6.6, "gute Einstellung", vNS, gesi, "%", ".90", ".40")
 HBC = LabWert(frm, "HBC", ">", 0.1, "Hepatitis B", "B16.-", gesi, vNS, "B16")
 HBSAG_K = LabWert(frm, "HBSAG_K", ">", 0.1, "Hepatitis B", "B16.-", gesi, vNS, "B16")
 HCV = LabWert(frm, "HCV", ">", 0, "(Z.n.)Hep.C", "B18.2", gesi, vNS, "B17", "B18")
 HCYP = LabWert(frm, "HCY-P", ">", 10.3, "Hyperhomozysteinämie", "E72.1", gesi, "µmol/l", "E72.1")
 HGEM_s = LabWert0(frm, "HGEM_S", "<", 20, "Ehrlichiose", "B99", gesi, vNS, "B99")
 HS = LabWert(frm, "HS", ">", 8.2, "Hyperurikämie", "E79.0", gesi, "mg/dl", "E79.0", "E79.0", 8.2, 6.1)
 Hsre = LabWert(frm, "HSRE", ">", 8.2, "Hyperurikämie", "E79.0", gesi, "mg/dl", "E79.0", "E79.0", 8.2, 6.1)
 MG_K = LabWert0(frm, "MG_K", "<", 0.66, "Magnesiummangel", "E61.2", gesi, "mmol/l", "E61.2")
 IA2AK = LabWert(frm, "IA2AK", ">", 0.1, "D.m.1 (IA2)", "E10.91", gesi, "U/ml", "E10")
 IA2AK = LabWert(frm, "IA2AK", "<", 0.1, "D.m.2 (IA2)", "E11.91", gesi, "U/ml", "E11", "R73")
 IAK_s = LabWert(frm, "IAK_S", ">", 0.1, "D.m.1 (IA)", "E10.91", gesi, "U/ml", "E10")
 IAK_s = LabWert(frm, "IAK_S", "<", 0.1, "D.m.2 (IA)", "E11.91", gesi, "U/ml", "E11", "R73")
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", Mitte Labor)"
#End If
 Na = LabWert0(frm, "NA", "<", 133, "Hyponatriämie", "E87.1", gesi, "mmol/l", "E87.1")
 Na = LabWert(frm, "NA", ">", 146, "Hypernatriämie", "E87.1", gesi, "mmol/l", "E87.0")
 IGE_K = LabWert(frm, "IGE_K", ">", 25, "Atopie", "T78.4", gesi, "U/ml", "T78", "T78", 25, 100)
 IGF1_S = LabWert(frm, "IGF-1_S", ">", 359, "Akromegalie", "E22.0", Va, "ng/ml", "E22", "E22", 316)
 INSELA = LabWert(frm, "INSELA", ">", 10, "D.m.1 (ICAa)", "E10.91", gesi, vNS, "E10")
 INSELA = LabWert(frm, "INSELA", "<", 10, "D.m.2 (ICAa)", "E11.91", gesi, vNS, "E11", "R73")
 INSELQ = LabWert(frm, "INSELQ", ">", 10, "D.m.1 (ICAq)", "E10.91", gesi, vNS, "E10")
 INSELQ = LabWert(frm, "INSELQ", "<", 10, "D.m.2 (ICAq)", "E11.91", gesi, vNS, "E11", "R73")
 k = LabWert(frm, "K", ">", 5.5, "Hyperkaliämie", "E87.5", gesi, "mmol/l", "E87.5", "E87.5", 5.5)
 k = LabWert0(frm, "K", "<", 3.6, "Hypokaliämie", "E87.6", gesi, "mmol/l", "E87.6", "E87.6", 3.6)
 LACT_E = LabWert(frm, "LACT_E", ">", 20, "Lactat-Azidose", "E87.2", gesi, "mg/dl", "E87.2")
 If Arq > 3 And Aldo_p > 15 Then _
  Call LabWert(frm, "ALDO_P", ">", 15, "Conn-Syndrom", "E26.0", Va, vNS, "E26.0")
 If Adrenu > 20 Or (Alter < 10 And Adrenu > 10) Then _
 Call LabWert(frm, "ADRENU", ">", 10, "Phäochromozytom", "D35.0", Va, vNS, "D35.0")
#If True Then
 Call LabwertA(frm, LA_LDL, ">", 130, "Hypercholesterinämie", "E78.0", gesi, "mg/dl", "E78", "E78", 160)
#Else
' LDL = LabWert(frm, "LDL", ">", 130, "Hypercholesterinämie", "E78.0", gesi, "mg/dl", "E78", "E78", 160)
' LDLB = LabWert(frm, "LDLB", ">", 130, "Hypercholesterinämie", "E78.0", gesi, "mg/dl", "E78", "E78", 160)
' LDLH = LabWert(frm, "LDLH", ">", 130, "Hypercholesterinämie", "E78.0", gesi, "mg/dl", "E78", "E78", 160)
' LDLC = LabWert(frm, "LDLC", ">", 130, "Hypercholesterinämie", "E78.0", gesi, "mg/dl", "E78", "E78", 160)
' LDLH01 = LabWert(frm, "LDLH01", ">", 130, "Hypercholesterinämie", "E78.0", gesi, "mg/dl", "E78", "E78", 160)
' LDLLG = LabWert(frm, "LDLLG", ">", 130, "Hypercholesterinämie", "E78.0", gesi, "mg/dl", "E78", "E78", 160)
' LDLS = LabWert(frm, "LDLS", ">", 130, "Hypercholesterinämie", "E78.0", gesi, "mg/dl", "E78", "E78", 160)
#End If
 Leuko = LabWert(frm, "LEUKO", ">", 10.5, "Entzündung", "D72.8", gesi, "/nl", "D72.8", "T8")
 LPa = LabWert(frm, "LP A", ">", 30, "Fettstoffwechselstörung", "E78.9", gesi, "mg/dl", "E78")
' ab hier muß weitergearbeitet werden
 MAK = LabWert(frm, "MAK", ">", 40, "Autoimmune SD-Krht", "E05.-", gesi, "IU/ml", "E05", "E06", 100)
 If (obWeib And Hb < 12.6) Or (Not obWeib And Hb < 14) Then
  MCH = LabWert(frm, "MCH", "<", 28, "Mikrozyt. Anämie", "D50.0", gesi, "pg", "D5", "D6", "K9")
  MCH = LabWert(frm, "MCH", ">", 34, "Makrozyt. Anämie", "D52.9", gesi, "pg", "D5", "D6", "K9")
 End If
 NH4 = LabWert(frm, "NH4", ">", 94, "Hepatische Enzephalopathie", "K72.7-", gesi, "µmol/l", "K72.7", "K74.6", 94, 82)
 P_ANCQ_S = LabWert(frm, "P_ANCQ_S", ">", 2, "mPAN/GN/RA/LE", "M31.-", Va, vNS, "M31", "N05")
 PANC_s = LabWert(frm, "P-ANC_S", ">", 2, "mPAN/GN/RA/LE", "M31.-", Va, vNS, "M31", "N05")
 PANELF = LabWert0(frm, "PANELF", "<", 300, "Pankreasinsuffizienz", "K86.8", gesi, "µg/g", "K86.8", "K86.1", 200)
 PCAK_cp = LabWert0(frm, "PC-AK_CP", "<", 70, "Protein-C-Mangel", "D68.8", Va, "%", "I82.9")
 PHOS = LabWert(frm, "PHOS", ">", 4.8, "sek. Hyperpara", "N25.8", gesi, "mg/dl", "N25.8", "E21")
 PO4 = LabWert(frm, "PO4", ">", 4.8, "sek. Hyperpara", "N25.8", gesi, "mg/dl", "N25.8", "E21")
 PROL_K = LabWert(frm, "PROL_K", ">", 619, "Prolaktinom", "D35.2", gesi, "mIU/l", "D35")
 pSa = LabWert(frm, "PSA", ">", 4, "Prostata-Ca", "C61", gesi, "ng/ml", "N40", "C61", 10)
 PSAK_cp = LabWert0(frm, "PS-AK_CP", "<", 70, "Protein-S-Mangel", "D68.8", Va, "%", "I82.9")
 rf = LabWert(frm, "RF", ">", 20, "Rheuma", "M05.-", gesi, "U/ml", "M05", "M06")
 RFQ = LabWert(frm, "RFQ", ">", 20, "Rheuma", "M05.-", gesi, "U/ml", "M05", "M06")
 ST24 = LabWert(frm, "ST 24", ">", 2, "Schwangerschaft", vNS, gesi, "IU/l", vNS, vNS, 10)
 THRM = LabWert0(frm, "THRM", "<", 130, "Thrombopenie", "D69.5-", gesi, "/nl", "D69", "K74", 50)
 THRM = LabWert(frm, "THRM", ">", 430, "Thrombozytose", "D75.2", gesi, "/nl", "D75", "D75", 1000)
 THROMB = LabWert0(frm, "THROMB", "<", 130, "Thrombopenie", "D69.5-", gesi, "/nl", "D69", "K74", 50)
 THROMB = LabWert(frm, "THROMB", ">", 430, "Thrombozytose", "D75.2", gesi, "/nl", "D75.9", "D75.9", 1000)
 TRAK_k = LabWert(frm, "TRAK_K", ">", 1, "M.Basedow", "E05.0", gesi, "U/l", "E05", "E05", 2)
 #If True Then
 Call LabwertA(frm, LA_TSH, "<", 0.3, "Hyperthyreose", "E05.9", Va, "µU/ml", "E05")
 Call LabwertA(frm, LA_TSH, ">", 4.2, "Hypothyreose", "E03.9", Va, "µU/ml", "E03", "E03", 10)
 Call LabwertA(frm, LA_TSH, ">", 6, "Hypothyreose", "E03.9", Va, "µU/ml", "E03", "E03", 10)
 #Else
' TSH = LabWert0(frm, "TSH", "<", 0.3, "Hyperthyreose", "E05.9", Va, "µU/ml", "E05")
' TSH = LabWert(frm, "TSH", ">", 4.5, "Hypothyreose", "E03.9", gesi, "µU/ml", "E03", "E03", 10)
' TSBF = LabWert0(frm, "TSBF", "<", 0.3, "Hyperthyreose", "E05.9", Va, "µU/ml", "E05")
' TSBF = LabWert(frm, "TSBF", ">", 4.5, "Hypothyreose", "E03.9", gesi, "µU/ml", "E03", "E03", 10)
' TSHL_k = LabWert0(frm, "TSH-L_K", "<", 0.3, "Hyperthyreose", "E03.9", Va, "µU/ml", "E05.9")
' TSHL_k = LabWert(frm, "TSH-L_K", ">", 4.5, "Hypothyreose", "E03.9", gesi, "µU/ml", "E03", "E03", 10)
 #End If
 VITD25_s = LabWert0(frm, "VITD25_S", "<", 10, "Vitamin-D-Mangel", "E55.9", gesi, "ng/ml", "E55", "E55")
 WACHSH_s = LabWert(frm, "WACHSH_S", ">", 7, "Akromegalie", "D35.2", Va, "ng/ml", "E22")
 syscmd 4, "Formularvorbereitung 8 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 
 Dim dm$
' Dim maxHbA1c#
 Dim rsDM As New ADODB.Recordset
 'sql = "SELECT IF(w1>w2 OR ISNULL(w2),w1,w2) wert FROM (SELECT n.pat_id,(SELECT MAX(CONVERT(REPLACE(CONCAT('0',trim(ln.wert)),',','.'),DECIMAL(9,2))) FROM laborneu ln WHERE ln.pat_id = n.pat_id AND ln.abkü RLIKE '^hba[1c]' AND CONVERT(REPLACE(CONCAT('0',trim(ln.wert)),',','.'),DECIMAL(9,2)) <22 AND ln.wert RLIKE '^[0-9 .,]*$') w1,(SELECT MAX(CONVERT(REPLACE(CONCAT('0',trim(w.wert)),',','.'),DECIMAL(9,2))) FROM laborywert w LEFT JOIN laboryus u ON w.usid = u.id WHERE u.pat_id = n.pat_id AND w.abkü RLIKE '^hba[1c]' AND CONVERT(REPLACE(CONCAT('0',trim(w.wert)),',','.'),DECIMAL(9,2))<22 AND w.wert RLIKE '^[0-9 .,]*$') w2 FROM namen n) i WHERE pat_id = " & Pat_id '  wg. falscher Fremdlabore jew. gestrichen: AND w.einheit = '%'
 sql = "SELECT wert FROM maxHbA1c WHERE pat_id=" & Pat_ID
 Set rsDM = Nothing
 myFrag rsDM, sql
 If Not rsDM.EOF Then
  If rsDM.Fields(0) >= 6.5 Then
   dm = "HbA1c " & rsDM.Fields(0)
  End If
 End If
 Dim maxGluc#
' sql = "SELECT IF(w1>w2 OR ISNULL(w2),w1,w2) wert FROM (SELECT n.pat_id,(SELECT MAX(CONVERT(REPLACE(CONCAT('0',trim(ln.wert)),',','.'),DECIMAL(9,2))) FROM laborneu ln WHERE ln.pat_id = n.pat_id AND ln.abkü RLIKE '^glu' AND ln.einheit = 'mg/dl' AND ln.wert RLIKE '^[0-9 .,]*$') w1,(SELECT MAX(CONVERT(REPLACE(CONCAT('0',trim(w.wert)),',','.'),DECIMAL(9,2))) FROM laborywert w LEFT JOIN laboryus u ON w.usid = u.id WHERE u.pat_id = n.pat_id AND w.abkü RLIKE '^glu' AND w.einheit = 'mg/dl'  AND w.wert RLIKE '^[0-9 .,]*$') w2 FROM namen n) i WHERE pat_id = " & Pat_id
 sql = "SELECT wert FROM maxGluc WHERE pat_id=" & Pat_ID
 Set rsDM = Nothing
 myFrag rsDM, sql
 If Not rsDM.EOF Then
  If rsDM.Fields(0) >= 200 Then
   dm = "Gluc " & rsDM.Fields(0)
  End If
 End If
 Call KRAdd(frm, dm, "Diabetes mellitus", "E11.91", gesi, "E10", "E11", "O24")
 
 
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 ccp = False
 If CREACL <> 0 Then
  Select Case Alter
   Case Is < 20
   Case Is < 30: CGr = IIf(obWeib, 72, 94)
   Case Is < 40: CGr = IIf(obWeib, 71, 89)
   Case Is < 50: CGr = IIf(obWeib, 50, 76)
   Case Is < 60: CGr = IIf(obWeib, 45, 54)
   Case Is < 70: CGr = IIf(obWeib, 37, 49)
   Case Is < 80: CGr = IIf(obWeib, 27, 30)
   Case Is < 90: CGr = IIf(obWeib, 26, 26)
  End Select
  If CREACL < CGr Then ccp = True
   If ccp Then _
    CREACL = LabWert(frm, "CREACL", "<", CGr, "Niereninsuffizienz", "N19", gesi, "ml/min", "N19", "N18", CGr, CGr - 10, "N17")
'  frm.MDi = frm.MDi + "Nierensinffizienz (Crea-Clearance " + CStr(CREACL) + " ml/min)" & vbcrlf
  End If
 End If 'CreaCL <> 0

 ' Blutdruck auswerten
' Dim rsRR AS DAO.Recordset
 Dim rsRR As New ADODB.Recordset
 Dim pzahl%
 pzahl = 0
' SET rsRR = TabÖff("RRParse", "ID")
' rsRR.Seek "=", Pat_id
' IF Not rsRR.NoMatch THEN
 frm.vTextB(150) = vNS
' rsRR.Open "SELECT * FROM rrparse WHERE pat_id = " & Pat_id & " ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
 myFrag rsRR, "SELECT * FROM rrparse WHERE pat_id = " & Pat_ID & " ORDER BY zeitpunkt DESC"
 If Not rsRR.BOF Then
  Do While Not rsRR.EOF
   If rsRR!RRsyst > 140 Or rsRR!RRdiast > 90 Then
    pzahl = pzahl + 1
   End If
   frm.vTextB(150) = frm.vTextB(150) & IIf(LenB(frm.vTextB(150)) = 0, vNS, ",") & rsRR!RRsyst & "/" & rsRR!RRdiast
   rsRR.MoveNext
  Loop
 End If ' not rsRR.bof THEN
 If pzahl = 1 Or pzahl = 2 Then
  Call KRAdd(frm, CStr(pzahl) + " Meßw.", "Arterielle Hypertonie", "I10.90", Va, "I10", , , , "vTextB", 150) ' RR
 ElseIf pzahl > 2 Then
  Call KRAdd(frm, CStr(pzahl) + " Meßw.", "Arterielle Hypertonie", "I10.90", gesi, "I10", , , , "vTextB", 150) ' R
 End If ' pzahl
 
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
' Dim bmi#
' IF frm.anaRS!Gewicht < 0 THEN frm.vTextB(10) = -frm.anaRS!Gewicht
' IF NOT ISNULL(frm.anaRS!Gewicht) AND NOT ISNULL(frm.anaRS!Größe) THEN
'  IF frm.anaRS!Größe > 0 THEN
'   bmi = frm.vTextB(10) * IIF(frm.vTextB(10) < 3, 100, 1) / frm.anaRS!Größe / frm.anaRS!Größe ' * 10000
'   Do
'    IF bmi = 0 THEN Exit Do
'    IF bmi > 8 THEN Exit Do
'    bmi = bmi * 10
'   Loop

'   IF frm.anaRS!bmi >= 28 THEN
'    Call KRAdd(frm, "BMI = " + Format$(frm.anaRS!bmi, "##,#"), "Übergewicht", "E66.99", gesi, "E66", , , , "vTextB", 11) ' Gewicht
'   END IF
   Select Case frm.anaRS!bmi
    Case Is >= 40
    Call KRAdd(frm, "BMI = " + Format$(frm.anaRS!bmi, "##,#"), "Übergewicht BMI>=40", "E66.92", gesi, "E66", , , , "vTextB", 11) ' Gewicht
    Case Is >= 35
    Call KRAdd(frm, "BMI = " + Format$(frm.anaRS!bmi, "##,#"), "Übergewicht BMI 35-<40", "E66.91", gesi, "E66", , , , "vTextB", 11) ' Gewicht
    Case Is >= 30
    Call KRAdd(frm, "BMI = " + Format$(frm.anaRS!bmi, "##,#"), "Übergewicht BMI 30-<35", "E66.90", gesi, "E66", , , , "vTextB", 11) ' Gewicht
   End Select
'  END IF
' END IF
 
 If InStrB(frm.vTextB(168), "Hypertonie") = 0 And Left$(frm.anaRS!Bluthochdruck, 1) = "j" Then ' MDi
  Call KRAdd(frm, "Ankreuzfeld", "Arterielle Hypertonie", "I10.90", gesi, "I10", , , , "vTextB", 58) ' Bluthochdruck
 End If
 
 Dim rpfl As Boolean
 Dim katfl%
 katfl = 0
 
 Dim AugFeld$(3), AugNr%(3)
 AugFeld(0) = "Augensp Befund"
 AugFeld(1) = "Weitere Anamnese"
 AugFeld(2) = "Sehminderung unbehebbar"
 AugFeld(3) = "Netzhaut gelasert"
 AugNr(0) = 66
 AugNr(1) = 102
 AugNr(2) = 68
 AugNr(3) = 67
 Dim i%
 Dim aktSt$
 
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", Netzhaut gelasert)"
#End If
 For i = 0 To UBound(AugFeld)
'  aktSt = LCase$(frm.anaRS.Fields(AugFeld(i)))
  If IsNull(frm.anaRS.Fields(AugFeld(i))) Then
   aktSt = vNS
  Else
   aktSt = LCase$(frm.anaRS.Fields(AugFeld(i)))
  End If
  If InStrB(aktSt, "glau") <> 0 Or (InStrB(aktSt, "grün") <> 0 And InStrB(aktSt, "star")) <> 0 Then
   Call KRAdd(frm, "Fragebogen", "Glaukom", "H40.9", gesi, "H40", , , , "vTextB", AugNr(i)) ' Augensp Befund
  End If
  If InStrB(aktSt, "atara") <> 0 Or (InStrB(aktSt, "grau") <> 0 And InStrB(aktSt, "star")) <> 0 Then
   Call KRAdd(frm, "Fragebogen", "Diab. Katarakt", "H28.0", gesi, "H28", "H26", , , "vTextB", AugNr(i)) ' Augensp Befund
   katfl = True
  End If
  rpfl = False
  If i = 0 Then
   If InStrB(aktSt, "keine") = 0 And InStrB(aktSt, "veränd") <> 0 Then rpfl = True
   If InStrB(aktSt, "op ") <> 0 Or InStrB(aktSt, "operier") <> 0 Then rpfl = True
   If InStrB(aktSt, "path") <> 0 And InStrB(aktSt, "path") <> InStrB(aktSt, "opath") + 2 Then rpfl = True
   If InStrB(aktSt, "blut") <> 0 Then rpfl = True
   If InStrB(aktSt, "schlecht") <> 0 Then rpfl = True
   If InStrB(aktSt, "o.B") <> 0 Or InStrB(aktSt, "o.b") <> 0 Or InStrB(aktSt, "oB") <> 0 Then rpfl = False
  End If
  If InStrB(aktSt, "retinop") <> 0 And InStrB(aktSt, "keine") = 0 Then
   rpfl = True
  End If
  If InStrB(aktSt, "gelasert") <> 0 Then rpfl = True
  If i = 4 And InStrB(LCase$(aktSt), "j") <> 0 Then
   If InStrB(aktSt, "wegen") = 0 And InStrB(aktSt, "wg") = 0 Then
      rpfl = True
   End If
  End If
  If rpfl Then
   Call KRAdd(frm, "Fragebogen", "Diab. Retinopathie", "H36.0", gesi, , , , , "vTextB", AugNr(i)) ' Augensp Befund
   If labPos <> 0 Then retflNeu = -1
   rpfl = True
  End If
 Next i
 If Not obkNeph And ((InStrB(LCase$(frm.anaRS![Diabet Nierenschaden]), "j") <> 0 Or frm.anaRS!Dialyse <> 0 Or diI("Z49.1", frm.anaRS!Pat_ID)) And InStrB(frm.vTextB(168), "Nephrop") = 0) Then ' MDi
  Call KRAdd(frm, "Fragebogen", "Diab.Nephropathie", "N08.3", gesi, , , , , "vTextB", 69) ' Diabet Nierenschaden
  If labPos <> 0 Then nieflNeu = -1
 End If
 If InStrB(LCase$(frm.anaRS![andere Nierenerkrankung]), "stein") <> 0 Then
  Call KRAdd(frm, "Fragebogen", "Nierensteinleiden", "N20.0", Zn, "N20", , , , "vTextB", 73) ' andere Nierenerkrankung
 End If
 
 Dim Anginap$, obKHK%, mi$, obMI%, ptca$, obPTCA%, hkBschr$
 obKHK = 0
 obMI = 0
 obPTCA = 0
 Anginap = IIf(IsNull(frm.anaRS![Angina pectoris]), vNS, LCase$(frm.anaRS![Angina pectoris]))
 If Left$(Anginap, 1) = "j" Then
  Call doGilbE(frm, "I25.9", "vTextB", 75) ' Angina pectoris
  obKHK = 1
 ElseIf InStrB(Anginap, "leicht") <> 0 Or InStrB(Anginap, "z.n.") <> 0 Or InStrB(Anginap, "evtl") <> 0 Or InStrB(Anginap, "manchmal") <> 0 Then
  Call doGilbE(frm, "I25.9", "vTextB", 75) ' Angina pectoris
  obKHK = 1
 End If
 mi = IIf(IsNull(frm.anaRS!Herzinfarkt), vNS, LCase$(frm.anaRS!Herzinfarkt))
 If Left$(mi, 1) = "j" Or InStrB(mi, "stummer") <> 0 Or InStrB(mi, "vwi") <> 0 Or InStrB(mi, "hwi") <> 0 Then
  Call doGilbE(frm, "I21.9", "vTextB", 76) ' Herzinfarkt
  obMI = 1
  obKHK = 2
 End If
 ptca = IIf(IsNull(frm.anaRS![PTCA oder Stent]), vNS, LCase$(frm.anaRS![PTCA oder Stent]))
 If InStrB(ptca, "j") <> 0 Then
  Call doGilbE(frm, "Z95.5Z", "vTextB", 78) ' PTCA oder Stent
  obPTCA = 1
  obKHK = 2
 End If
 If frm.anaRS![Bypaß kardial] <> 0 Then
  Call doGilbE(frm, "Z95.1", "vCheckB", 3) ' Bypaß kardial
  obKHK = 2
 End If
 hkBschr = IIf(IsNull(frm.anaRS![Herzkrankheit Beschreibung]), vNS, LCase$(frm.anaRS![Herzkrankheit Beschreibung]))
 If InStrB(hkBschr, "a.p.") <> 0 Or InStrB(hkBschr, "retrost") <> 0 Or InStrB(hkBschr, "-gef") <> 0 Or InStrB(hkBschr, "brustschmerz") <> 0 Then
  Call doGilbE(frm, "I25.9", "vTextB", 75) ' Angina pectoris
  obKHK = 2
 End If
 If obKHK = 2 Then
   If InStrB(hkBschr, "3-gef") <> 0 Or InStrB(hkBschr, "4-gef") <> 0 Or InStrB(hkBschr, "3-fach") <> 0 Or InStrB(hkBschr, "4-fach") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Koronare Dreigefäßerkrankung", "I25.13", gesi, "I25")
   ElseIf InStrB(hkBschr, "2-gef") <> 0 Or InStrB(hkBschr, "2-fach") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Koronare Zweigefäßerkrankung", "I25.12", gesi, "I25")
   ElseIf InStrB(hkBschr, "1-gef") <> 0 Or InStrB(hkBschr, "1-fach") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Koronare Eingefäßerkrankung", "I25.11", gesi, "I25")
   Else
    Call KRAdd(frm, "Fragebogen", "Koronare Herzerkrankung", "I25.9", gesi, "I25")
   End If
 ElseIf obKHK = 1 Then
  Call KRAdd(frm, "Fragebogen", "koronare Herzerkrankung", "I25.9", Va, "I25")
 End If
 If obPTCA = 1 Then
  Call KRAdd(frm, "Fragebogen", "PTCA oder Stent", "Z95.5", Zn, "Z95.5")
 End If
 If frm.anaRS![Bypaß kardial] <> 0 Then Call KRAdd(frm, "Fragebogen", "Bypaß kardial", "Z95.1", Zn)
 If obMI = 1 Then Call KRAdd(frm, "Fragebogen", "Myokardinfarkt", "I21.9", Zn, "I21", "I25.2")
 
 Dim hdb$
 hdb = IIf(IsNull(frm.anaRS!Hirndurchblutungsstörung), vNS, LCase$(frm.anaRS!Hirndurchblutungsstörung))
 If InStrB(hdb, "j") <> 0 Or InStrB(hdb, "z.n.") <> 0 Then
  Call KRAdd(frm, "Fragebogen", "Atheromatose hirnzuf. Gefäße", "I65.9", Va, "I65", "I66", "I68", "vTextB", 82) ' Hirndurchblutungsstörung")
  If labPos <> 0 Then angflNeu = -1
 End If
 Dim Schlag$
 Schlag = IIf(IsNull(frm.anaRS!Schlaganfall), vNS, LCase$(frm.anaRS!Schlaganfall))
 If Len(Schlag) > 0 And Schlag <> "-" And Schlag <> "n" Then
  If InStrB(Schlag, "gedächtnis") = 0 And InStrB(Schlag, "vergeßlich") = 0 And InStrB(Schlag, "vergesslich") = 0 And InStrB(Schlag, "entfällt") = 0 Then
   If InStrB(Schlag, "blutu") <> 0 Then
    Call KRAdd(frm, "Fragebogen", "Hirnblutung", "I61.9", Zn, "I61", "I62", , , "vTextB", 83) ' Schlaganfall
   Else
    Call KRAdd(frm, "Fragebogen", "Apoplexie", "I63.9", Zn, "I64", "I63", , , "vTextB", 83) ' Schlaganfall
   End If
  End If
 End If
 
 syscmd 4, "Formularvorbereitung 10 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 Dim BDBSt%, SFKht%
 BDBSt = 0
 SFKht = 0
 If Left$(LCase$(frm.anaRS!Beindurchblutungsstörung), 1) = "j" Then
  BDBSt = -1
  Call doGilb(frm, "I73.9", "vTextB", 84) ' Beindurchblutungsstörung
 End If
 If Left$(LCase$(frm.anaRS!Schaufensterkrankheit), 1) = "j" Then
  SFKht = -1
  Call doGilb(frm, "I73.9", "vTextB", 85) ' Schaufensterkrankheit
 End If
 If (BDBSt Or SFKht) And Not (BDBSt And Not SFKht And Not PStatPath And frm.anaRS![bypaß peripher] = 0 And IsNull(frm.anaRS!Geschwür) And IsNull(frm.anaRS![Amputation]) And IsNull(frm.anaRS![pAVK Beschreibung])) Then
  Call KRAdd(frm, "Fragebogen", "Periphere arterielle Verschlußkrankheit", "I73.9", gesi, "I73", "I74", "I70")
  If labPos <> 0 Then angflNeu = -1
 End If
 If frm.anaRS![bypaß peripher] <> 0 Then
  Call KRAdd(frm, "Fragebogen", "Bypaß peripher", "Z95.88", Zn, "I73", "I74", "I70", "Z95", "vCheckB", 4) ' bypaß peripher
  If labPos <> 0 Then angflNeu = -1
 End If
 Dim amput$
 amput = IIf(IsNull(frm.anaRS![Amputation]), vNS, LCase$(frm.anaRS![Amputation]))
 If Len(amput) > 1 And amput <> "-" And amput <> "n" And amput <> "entfällt" Then
  Call KRAdd(frm, "Fragebogen", "Amputation", "Z44.1", Zn, "Z44", "Z97", , , "vTextB", 87) ' Amputation
 End If
#If mitab Then
 Call DiabetesDiagnose(frm, pnpflNeu, KZahlNeu, anpflNeu, nieflNeu, angflNeu, bgwflNeu, katfl, retflNeu, nasflNeu, dfsflNeu, oblhNeu)
#End If
syscmd 4, "Formularvorbereitung 11 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", vor knöpfeanpassen)"
#End If
Call knöpfeanpassen(frm)
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", nach knöpfeanpassen)"
#End If
 
syscmd 4, "Formularvorbereitung 12 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
'Call do_Hirndurchblutungsstörung_Exit(0, frm)
 Dim rsVerSi As New ADODB.Recordset, rsv1 As New ADODB.Recordset, rsV0 As New ADODB.Recordset
 Dim rsVK As New ADODB.Recordset
 Dim Versicherung$
 sql = "SELECT * FROM `faelle` INNER JOIN (SELECT MIN(pat_id) AS pid, MAX(bhfb) AS bb FROM `faelle` GROUP BY pat_id) AS sel ON (`faelle`.`bhfb`=sel.bb) AND (`faelle`.`pat_id`=sel.pid)"
' SET rVK = Dtb.OpenRecordset(sql)
 myFrag rsVK, sql & " WHERE pat_id = " & Pat_ID
' rVK.FindFirst ("pat_id = " + CStr(Pat_id))
 Versicherung = vNS
 If Not rsVK.BOF Then
  Versicherung = rsVK!VKNr
 End If
' Versicherung = IIF(ISNULL(rVK!VKNr) OR rVK.NoMatch, vns, rVK!VKNr)
''If Not versi.BOF THEN frm.Versicherung = versi!Versicherung
''versi.Close
'' Aufrufe der Anamnese usw. belegen
 Call getDokPfad
 Dim rDok As ADODB.Recordset
 
 syscmd 4, "Formularvorbereitung 13 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 An1Pfad = PfadFestLeg("An1Aufruf", "%anamnese%1%", frm)
 An2Pfad = PfadFestLeg("An2Aufruf", "%anamnese%2%", frm)
 AnAPfad = PfadFestLeg("AnAAufruf", "%anamnese%allg%", frm, "%allg%anamnese%")
 CheckPfad = PfadFestLeg("CheckAufruf", "%che%kliste%", frm)
 
 Dim lz%, lzu%, rsLab As New ADODB.Recordset
 syscmd 4, "Formularvorbereitung 13a " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 'lz = Dtb.OpenRecordset("SELECT COUNT(0) AS ct FROM `" + QMdbAkt + "`.LaborDokumente WHERE pat_id = " + CStr(frm.Pat_id))!ct
 sql = "SELECT COUNT(0) AS ct " & _
 "FROM (briefe dl LEFT JOIN `br_abgehakt` da ON dl.Pfad=da.DokPfad) " & _
 "WHERE Name LIKE '%labor%' AND dl.pat_id = " & CStr(Pat_ID)
 Set rsLab = Nothing
 myFrag rsLab, sql
 If Not rsLab.BOF Then lz = rsLab!ct
 syscmd 4, "Formularvorbereitung 13b " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
' lzu = Dtb.OpenRecordset("SELECT COUNT(0) AS ct FROM `" + QMdbAkt + "`.LaborDokumente WHERE pat_id = " + CStr(frm.Pat_id) + " AND abgehakt")!ct
 sql = "SELECT COUNT(0) AS ct " & _
 "FROM (briefe dl LEFT JOIN `br_abgehakt` da ON dl.Pfad=da.DokPfad) " & _
 "WHERE Name LIKE '%labor%' AND abgehakt AND dl.pat_id = " & CStr(Pat_ID)
 Set rsLab = Nothing
 myFrag rsLab, sql
 If Not rsLab.BOF Then lzu = rsLab!ct
 syscmd 4, "Formularvorbereitung 13c " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 frm.vCommandB(6).Caption = CStr(lzu) + "&/" + CStr(lz) & " " & "LD" ' Labor
 frm.vCommandB(6).Enabled = True ' Labor
 frm.vCommandB(6).MaskColor = 0   ' Labor ' eigentlich ForeColor, gibt es aber nicht
 If lzu = lz Then
  If lz = 0 Then
   frm.vCommandB(6).Enabled = False ' Labor
  Else
   frm.vCommandB(6).MaskColor = RGB(100, 100, 100) ' Labor
  End If
 Else
 End If
 If 1 = 0 And lzu = 0 Then
   frm.vCommandB(6).Enabled = 0 ' Labor
 End If
' jetzt das gleiche für Briefe
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 syscmd 4, "Formularvorbereitung 13d " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 sql = "SELECT COUNT(0) AS ct FROM `briefe` WHERE pat_id = " + CStr(Pat_ID)
 Set rsLab = Nothing
 myFrag rsLab, sql
 If Not rsLab.BOF Then
  lz = rsLab!ct
  frm.vCommandB(9).Caption = CStr(lz) + " Briefe (7)" '(&7)
 End If
 sql = "SELECT MAX(zeitpunkt) AS zp FROM `briefe` WHERE (name LIKE 'Brief an Dr%' OR name LIKE 'Brief an F%Dr.' OR name LIKE '%Arztbrief%') AND NOT name LIKE '%Entwurf%' AND pat_id = " + CStr(Pat_ID)
 '!zp, "dd/mm/yy")
 Set rsLab = Nothing
 myFrag rsLab, sql
 If Not rsLab.BOF Then
  frm.vCommandB(9).Caption = frm.vCommandB(9).Caption & ", zul." & rsLab!Zp ' Briefe
 End If
' frm.briefe.Caption = CStr(lz) + " Briefe (&7), zul." + format$(Dtb.OpenRecordset("SELECT MAX(zeitpunkt) AS zp FROM `" + QMdbAkt + "`.`briefe` WHERE (name LIKE ""Brief an Dr*"" OR name LIKE ""Brief an F*Dr."" OR name LIKE ""*Arztbrief*"") AND NOT name LIKE ""*Entwurf*"" AND pat_id = " + CStr(Pat_id))!zp, "dd/mm/yy")
 frm.vCommandB(9).Enabled = True ' Briefe
 frm.vCommandB(9).MaskColor = 0 ' Briefe
 If lz = 0 Then
   frm.vCommandB(9).Enabled = 0 ' Briefe
 End If
 syscmd 4, "Formularvorbereitung 14 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname

' und für Blutzuckerkurven
 Dim maxzp As Date
 maxzp = 0
 sql = "SELECT COUNT(0) ct, MAX(zeitpunkt) mzp FROM briefe WHERE name LIKE '%BZ%' AND pat_id = " & CStr(Pat_ID)
 Set rsLab = Nothing
 myFrag rsLab, sql
 If Not rsLab.BOF Then
  lz = rsLab!ct
  If lz > 0 Then
   maxzp = rsLab!mzp
  End If
 End If
' BZKurven
 frm.vCommandB(11).Caption = CStr(lz) + " BZKur.(8), zul." & Format$(maxzp, "dd/mm/yy") ' BZKurven ' (&8)
 frm.vCommandB(11).Enabled = True ' BZKurven
 frm.vCommandB(11).MaskColor = 0  ' BZKurven
 If lz = 0 Then
   frm.vCommandB(11).Enabled = 0 ' BZKurven
 End If
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
' und für Augenbefunde
' lz = Dtb.OpenRecordset("SELECT COUNT(0) AS ct FROM `" + QMdbAkt + "`.`dokumente` WHERE (dokname LIKE ""*augenb`!l`*"" OR dokname LIKE ""*augen`aä`rzt*"" OR dokname LIKE ""*aa"") AND pat_id = " + CStr(frm.Pat_id))!ct + _
      Dtb.OpenRecordset("SELECT COUNT(0) AS ct FROM `" + QMdbAkt + "`.`eintraege` WHERE (inhalt LIKE ""*augenb`!l`*"" OR inhalt LIKE ""*augen`aä`rzt*"" OR inhalt LIKE ""* aa`!g`*"") AND pat_id = " + CStr(Pat_id))!ct
 lz = 0
 maxzp = 0
 Set rsLab = Nothing
 sql = "SELECT COUNT(0) ct,MAX(zeitpunkt) mzp FROM briefe b WHERE ((name LIKE '%augenb%' AND NOT name LIKE '%augenbl%') OR name LIKE '%augenarzt%' OR name LIKE '%augenärzt%' OR name LIKE '%aa') AND pat_id = " & CStr(Pat_ID) ' name REGEXP '.*augenb[^l].*'
 Set rsLab = Nothing
 myFrag rsLab, sql
 If Not rsLab.BOF Then
  lz = rsLab!ct
  If lz > 0 Then
   maxzp = rsLab!mzp
  End If
 End If
 sql = "SELECT COUNT(0) AS ct,MAX(zeitpunkt) AS mzp FROM `eintraege` WHERE ((inhalt LIKE '%augenb%' AND NOT inhalt LIKE '%augenbl%') OR inhalt LIKE '%augenarzt%' OR inhalt LIKE '%augenärzt%' OR (inhalt LIKE '% aa' AND NOT inhalt LIKE '% aag%')) AND pat_id = " + CStr(Pat_ID)
 Set rsLab = Nothing
 myFrag rsLab, sql
 If Not rsLab.BOF Then
  lz = lz + rsLab!ct
  If rsLab!mzp > maxzp Then maxzp = rsLab!mzp
 End If
 
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 frm.vCommandB(13).Caption = CStr(lz) + " Aug'b.(9), zul." & Format$(maxzp, "dd/mm/yy") ' AugenBefunde ' (&9)
 frm.vCommandB(13).Enabled = True ' AugenBefunde
 frm.vCommandB(13).MaskColor = 0 ' AugenBefunde
 If lz = 0 Then
   frm.vCommandB(13).Enabled = 0 ' AugenBefunde
 End If
 
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ")"
#End If
 Dim rMA As New ADODB.Recordset, zp0 As Date, zp1 As Date, zp2 As Date, zp3 As Date
 myFrag rMA, "SELECT COUNT(0) ct, COALESCE(MAX(zeitpunkt),0) mzp FROM `eintraege` WHERE art IN ('tk','ARCHIE2','APK') AND pat_id = " & CStr(Pat_ID)
 frm.MA0Zahl = rMA!ct ' Kothny
 zp0 = rMA!mzp
 Set rMA = Nothing
 myFrag rMA, "SELECT COUNT(0) ct, COALESCE(MAX(zeitpunkt),0) mzp FROM `eintraege` WHERE (art IN ('gs','doppler','duplex') OR inhalt LIKE '%(gs)%') AND pat_id = " & CStr(Pat_ID)
 frm.MA1Zahl = rMA!ct ' Schade
 zp1 = rMA!mzp
 Set rMA = Nothing
 myFrag rMA, "SELECT COUNT(0) ct, COALESCE(MAX(zeitpunkt),0) mzp FROM `eintraege` WHERE (art IN ('wd') OR inhalt LIKE '%(wd)%') AND pat_id = " & CStr(Pat_ID)
 frm.MA2Zahl = rMA!ct ' Wagner
 zp2 = rMA!mzp
 Set rMA = Nothing
 myFrag rMA, "SELECT COUNT(0) ct, COALESCE(MAX(zeitpunkt),0) mzp FROM `eintraege` WHERE (art IN ('ah') OR inhalt LIKE '%(ah)%') AND pat_id = " & CStr(Pat_ID)
 frm.MA3Zahl = rMA!ct ' Wagner
 zp3 = rMA!mzp
 Set rMA = Nothing
 If zp0 > zp1 And zp0 > zp2 Then
  frm.MA0Lab.FontBold = True
  frm.MA1Lab.FontBold = False
  frm.MA2Lab.FontBold = False
 ElseIf zp1 > zp0 And zp1 > zp2 Then
  frm.MA0Lab.FontBold = False
  frm.MA1Lab.FontBold = True
  frm.MA2Lab.FontBold = False
 ElseIf zp2 > zp1 And zp2 > zp0 Then
  frm.MA0Lab.FontBold = False
  frm.MA1Lab.FontBold = False
  frm.MA2Lab.FontBold = True
 Else
  frm.MA0Lab.FontBold = False
  frm.MA1Lab.FontBold = False
  frm.MA2Lab.FontBold = False
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Form_Current2/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Form_Current2
#End If ' mitab

#If mitab Then
 Function DiabetesDiagnose(frm As AnBog, pnpflNeu%, KZahlNeu%, anpflNeu%, nieflNeu%, angflNeu%, bgwflNeu%, katfl%, retflNeu%, nasflNeu%, dfsflNeu%, oblhNeu%)
'Diabetesdiagnose
 Dim DICD$, DICDg$, DICDh$, DICDj$
 Dim KZahl%, anpfl As Boolean, pnpfl As Boolean, retfl As Boolean, angfl As Boolean, dfsfl As Boolean, _
 nasfl As Boolean, bgwfl As Boolean, niefl As Boolean, ketfl As Boolean, hypfl As Boolean
 Dim DiabetesName$
 DiabetesName = "Diabetes mellitus"
 On Error GoTo fehler
 Select Case frm.anaRS!Diabetestyp
  Case "1"
   DICD = "E10"
  Case "2"
   DICD = "E11"
  Case "-"
   DICD = "-"
  Case "?"
   DICD = "E14"
  Case "s"
   DICD = "E13"
  Case "g"
   DICD = "O24.4"
   DiabetesName = "Gestationsdiabetes"
  Case "p"
   DICD = "R73.0"
   DiabetesName = "Pathologische Glucosetoleranz"
 End Select
 KZahl = 0
 pnpfl = 0
 anpfl = 0
 retfl = 0
 nasfl = 0
 bgwfl = 0
 niefl = 0
 ketfl = 0
 hypfl = 0
 angfl = 0
 dfsfl = 0
 niefl = False
 If Not IsNull(frm.vTextB(147)) Then
 If InStrB(frm.vTextB(147), "G63") <> 0 Or InStrB(frm.vTextB(147), "G62") <> 0 Then
  pnpfl = True
  KZahl = KZahl + 1
 ElseIf pnpflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 If InStrB(frm.vTextB(147), "G99") <> 0 Then
  anpfl = True
  KZahl = KZahl + 1
 ElseIf anpflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 If Not obkNeph Then
  If (InStrB(frm.vTextB(147), "N1") <> 0 Or InStrB(frm.vTextB(147), "N08.3") <> 0) Then
   niefl = True
   KZahl = KZahl + 1
  ElseIf nieflNeu Then
   KZahlNeu = KZahlNeu + 1
  End If
 End If
 If InStrB(frm.vTextB(147), "I7") <> 0 Or InStrB(frm.vTextB(147), "I6") <> 0 Then
  angfl = True
  KZahl = KZahl + 1
 ElseIf angflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 If InStrB(frm.vTextB(147), "M14.2") <> 0 Then
  bgwfl = True
  KZahl = KZahl + 1
 ElseIf bgwflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 If InStrB(frm.vTextB(147), "H28.0") <> 0 Then
  katfl = True
  KZahl = KZahl + 1
 End If
 If InStrB(frm.vTextB(147), "H36.0") <> 0 Then
  retfl = True
  KZahl = KZahl + 1
 ElseIf retflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 If InStrB(frm.vTextB(147), "K71") <> 0 Then
  nasfl = True
  KZahl = KZahl + 1
 ElseIf nasflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 If InStrB(frm.vTextB(147), "E87.2") <> 0 Then
  ketfl = True
  KZahl = KZahl + 1
 End If
 If InStrB(frm.vTextB(147), "E16") <> 0 Then
  hypfl = True
  KZahl = KZahl + 1
 End If
 If InStrB(frm.vTextB(147), "L89") <> 0 Then
  dfsfl = True
  KZahl = KZahl + 1
 ElseIf dfsflNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 
 Dim oblhAlt%
 oblhAlt = 0
 If InStrB(frm.vTextB(147), "E88") <> 0 Or InStrB(frm.vTextB(147), "L94") <> 0 Then
  oblhAlt = True
  KZahl = KZahl + 1
 ElseIf oblhNeu Then
  KZahlNeu = KZahlNeu + 1
 End If
 End If
 
 syscmd 4, "Formularvorbereitung 9 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 DICDg = DICD
 DICDj = vNS
 If Left$(DICD, 2) = "E1" Then
  DICDg = DICDg + "."
  Select Case KZahl
   Case Is > 1
    DICDg = DICDg + "7"
   Case 1
    If ketfl Then
     DICDg = DICDg + "1"
    ElseIf niefl Then
     DICDg = DICDg + "2"
    ElseIf retfl Or katfl Then
     DICDg = DICDg + "3"
    ElseIf anpfl Or pnpfl Then
     DICDg = DICDg + "4"
    ElseIf angfl Then
     DICDg = DICDg + "5"
    ElseIf nasfl Or hypfl Or oblhAlt Or dfsfl Then
     DICDg = DICDg + "6"
    End If
   Case 0
    DICDg = DICDg + "9"
  End Select
  Select Case KZahl
   Case Is > 1
    If dfsfl Then
     DICDg = DICDg & "5"
    Else
     DICDj = DICDg & "5"
     DICDg = DICDg & "3"
    End If
   Case Else
    DICDg = DICDg + "1"
  End Select
 End If
 If DICDg <> "-" Then Call KRAdd(frm, DICDg, DiabetesName, DICDg, gesi, "E1", "O24.4", "R73.0")

 syscmd 4, "Formularvorbereitung 10 " & frm.anaRS!Nachname & ", " & frm.anaRS!Vorname
 DICDh = DICD
 If Left$(DICD, 2) = "E1" Then
  DICDh = DICDh + "."
  Select Case KZahl + KZahlNeu
   Case Is > 1
    DICDh = DICDh + "7"
   Case 1
    If ketfl Then
     DICDh = DICDh + "1"
    ElseIf niefl Or nieflNeu Then
     DICDh = DICDh + "2"
    ElseIf retfl Or retflNeu Or katfl Then
     DICDh = DICDh + "3"
    ElseIf anpfl Or pnpfl Or anpflNeu Or pnpflNeu Then
     DICDh = DICDh + "4"
    ElseIf angfl Or angflNeu Then
     DICDh = DICDh + "5"
    ElseIf nasfl Or nasflNeu Or hypfl Or oblhAlt Or oblhNeu Or dfsfl Or dfsflNeu Then
     DICDh = DICDh + "6"
    End If
   Case 0
    DICDh = DICDh + "9"
  End Select
  Select Case KZahl + KZahlNeu
   Case Is > 1
    If dfsfl Then
     DICDh = DICDh & "5"
    Else
     If LenB(DICDj) = 0 Then DICDj = DICDh & "5"
     DICDh = DICDh & "3"
    End If
   Case Else
    DICDh = DICDh & "1"
  End Select
 End If
 If DICDh <> "-" And DICDg <> DICDh Then Call KRAdd(frm, DICDh, DiabetesName, DICDh, gesi, DICDh, "O24.4", "R73.0")
 If DICDj <> "" And DICDj <> DICDg And DICDj <> DICDh Then Call KRAdd(frm, DICDj, DiabetesName, DICDj, gesi, DICDj, "O24.4", "R73.0")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DiabetesDiagnose/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Diabetesdiagnose

Function ddsono(frm As AnBog)
Dim rEintr As New ADODB.Recordset, i&, aktart$
For i = 177 To 180
  frm.vTextB(i) = vNS
  Select Case i
   Case 177: aktart = "'doppler','duplex'"
   Case 178: aktart = "'anal','andm','andm2'"
   Case 179: aktart = "'sono','sd'"
   Case 180: aktart = artSpezEintr
  End Select
#If False Then
  myFrag rEintr, "SELECT * FROM `eintraege` WHERE pat_id = " & CStr(Pat_ID) & " AND art IN (" & aktart & ") ORDER BY zeitpunkt DESC"
  If Not rEintr.BOF Then
   frm.vTextB(i) = frm.vTextB(i) & rEintr!art + " vom " + Format$(rEintr!Zeitpunkt, "dd.mm.yy:") ' `eintraege`
   Do While Not rEintr.EOF
    frm.vTextB(i) = frm.vTextB(i) & rEintr!art + ": " + rEintr!Inhalt + vbCrLf ' `eintraege`
    rEintr.Move 1
   Loop
  End If
#Else
  myFrag rEintr, "SELECT REPLACE(GROUP_CONCAT(CONCAT(DATE_FORMAT(zeitpunkt,'%d.%m.%y'),'(',art,'): ',inhalt,CHAR(13),CHAR(10))),CONCAT(CHAR(13),CHAR(10),','),CONCAT(CHAR(13),CHAR(10))) i FROM `eintraege` WHERE pat_id = " & CStr(Pat_ID) & " AND art IN (" & aktart & ") ORDER BY zeitpunkt DESC"
  If Not IsNull(rEintr!i) Then
   frm.vTextB(i) = rEintr!i
  End If
#End If
  Set rEintr = Nothing
Next i
'  frm.vTextB(180) = vNS ' eintraege
'  frm.vTextB(177) = vNS ' Doppler, Duplex
'  frm.vTextB(178) = vNS ' Anamnese
'  frm.vTextB(179) = vNS ' Sono = vns
'  myFrag rEintr, "SELECT * FROM `eintraege` WHERE pat_id = " + CStr(Pat_id) + " AND art IN (" & artSpezEintr & ") ORDER BY zeitpunkt DESC"
'  IF Not rEintr.BOF THEN
'   frm.vTextB(180) = rEintr!Art + " vom " + Format$(rEintr!Zeitpunkt, "dd.mm.yy:") ' `eintraege`
'   Do While Not rEintr.EOF
'    frm.vTextB(180) = frm.vTextB(180) & rEintr!Art + ": " + rEintr!Inhalt + vbCrLf ' `eintraege`
'    rEintr.Move 1
'   Loop
'  END IF
'  SET rEintr = Nothing
'  myFrag rEintr, "SELECT * FROM `eintraege` WHERE pat_id = " + CStr(Pat_id) + " AND art IN ('doppler','duplex') ORDER BY zeitpunkt DESC"
'  IF Not rEintr.BOF THEN
'   frm.vTextB(177) = frm.vTextB(177) & rEintr!Art + " vom " + Format$(rEintr!Zeitpunkt, "dd.mm.yy:")
'   Do While Not rEintr.EOF
'    frm.vTextB(177) = frm.vTextB(177) & rEintr!Art + ": " + rEintr!Inhalt + vbCrLf
'    rEintr.Move 1
'   Loop
'  END IF
'  SET rEintr = Nothing
'  myFrag rEintr, "SELECT * FROM `eintraege` WHERE pat_id = " + CStr(Pat_id) + " AND art IN ('anal','andm') ORDER BY art, zeitpunkt DESC"
'  IF Not rEintr.BOF THEN
'   frm.vTextB(178) = frm.vTextB(178) & rEintr!Art + " vom " + Format$(rEintr!Zeitpunkt, "dd.mm.yy:")
'   Do While Not rEintr.EOF
'    frm.vTextB(178) = frm.vTextB(178) & rEintr!Art + ": " + rEintr!Inhalt + vbCrLf
'    rEintr.Move 1
'   Loop
'  END IF
'  SET rEintr = Nothing
'  myFrag rEintr, "SELECT * FROM `eintraege` WHERE pat_id = " + CStr(Pat_id) + " AND art IN (""sono"",""sd"") ORDER BY zeitpunkt DESC"
'  IF Not rEintr.BOF THEN
'   frm.vTextB(179) = frm.vTextB(179) & rEintr!Art + " vom " + Format$(rEintr!Zeitpunkt, "dd.mm.yy:")
'   Do While Not rEintr.EOF
'    frm.vTextB(179) = frm.vTextB(179) & rEintr!Art + ": " + rEintr!Inhalt + vbCrLf
'    rEintr.Move 1
'   Loop
'  END IF
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ddsono/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ddsono

Function sensib(Pat_ID&, Optional ByRef etext) As DMPStat
  Dim sql$, obpath&, pz&, gz&, erg$
  Dim rs As New ADODB.Recordset
  On Error GoTo fehler
'#Const testen = True
#If testen Then
  Dim rf As New ADODB.Recordset
  Lese.ProgStart
  myFrag rf, "SELECT pat_id, gesname(pat_id) name FROM aktfv"
  Do While Not rf.EOF
   myEFrag ("CALL sensib(" & rf!Pat_ID & ",@sens_obpath,@sens_pz, @sens_gz, @sens_text)")
   myFrag rs, "SELECT @sens_obpath,@sens_pz,@sens_gz, @sens_text"
'   Debug.Print rf!Pat_id, rs![@sens_obpath], rs![@sens_pz], rs![@sens_gz], rs![@sens_text] ' rf!name,
   Set rs = Nothing
   rf.MoveNext
  Loop
#Else
  myEFrag ("CALL sensib(" & Pat_ID & ",@sens_obpath,@sens_pz, @sens_gz, @sens_text)")
  myFrag rs, "SELECT @sens_obpath, @sens_text"
  On Error Resume Next
  etext = rs![@sens_text]
  sensib = rs![@sens_obpath]
#End If
' So ging's nicht mit dem String, lieferte mit allen möglichen Versuchen nur ?????
'  Dim rCmd As New ADODB.Command
'  rCmd.ActiveConnection = DBCn
'  rCmd.CommandType = adCmdStoredProc
'  rCmd.CommandText = "sensib"
'  rCmd.Parameters.Append rCmd.CreateParameter("pid", adInteger, adParamInput, 6, pat_id)
'  rCmd.Parameters.Append rCmd.CreateParameter("obpath", adInteger, adParamOutput, 6, obpath)
'  rCmd.Parameters.Append rCmd.CreateParameter("ergeb", adVarChar, adParamOutput, 60, erg) ', 60, erg)
'  SET rs = rCmd.Execute
'  IF Not rs.EOF THEN
'   Debug.Print rs.Fields(0), rs!obpath, rs!ergeb
'  END IF
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in sensib/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' sensib

 ' 0 = unauffällig, 1 = nicht untersucht, 2 = auffällig
 Function PStatNeu(Pat_ID&, Optional ByRef etext) As DMPStat
  Dim sql$, rpuls As New ADODB.Recordset
  On Error GoTo fehler
'Call Lese.ProgStart
  sql = "SELECT n.pat_id, " & _
        "IF (ISNULL(e.inhalt) AND ISNULL(u.pat_id) AND ISNULL(icd),1, " & _
        " IF (NOT ISNULL(e.inhalt) AND ((e.inhalt LIKE '%pst%' OR e.inhalt LIKE '%postst%') OR ISNULL(icd)), " & _
        "      IF(e.inhalt LIKE '%pst%' OR e.inhalt LIKE '%postst%',2,0), " & _
        "  IF (ISNULL(u.pat_id), " & _
        "   IF(ISNULL(icd),1,2), " & _
        "    IF((INSTR(pulsatp_re,'+')<>0 AND INSTR(pulsatp_li,'+')<>0) OR (INSTR(pulsadp_re,'+')<>0 AND INSTR(pulsadp_li,'+')<>0),0,2) " & _
        "))) ergeb, " & _
        "IF (ISNULL(e.inhalt) AND ISNULL(u.pat_id) AND ISNULL(icd),'nicht untersucht', " & _
        " IF (NOT ISNULL(e.inhalt) AND ((e.inhalt LIKE '%pst%' OR e.inhalt LIKE '%postst%') OR ISNULL(icd)), " & _
        "      CONCAT(IF(e.inhalt LIKE '%pst%' OR e.inhalt LIKE '%postst%','','un'),'auffällig (',upper(LEFT(e.art,1)),MID(e.art,2),' ',DATE_FORMAT(e.zeitpunkt,'%e.%c.%Y'),')'), " & _
        "  IF (ISNULL(u.pat_id), " & _
        "   IF(ISNULL(icd),'nicht untersucht',CONCAT('auffällig (Diagn.',icd,')')), " & _
        "    CONCAT(IF((INSTR(pulsatp_re,'+')<>0 AND INSTR(pulsatp_li,'+')<>0) OR (INSTR(pulsadp_re,'+')<>0 AND INSTR(pulsadp_li,'+')<>0),'un',''),'auffällig: ',pulsatp_re,'/',pulsatp_li,',',pulsadp_re,'/',pulsadp_li,' (',DATE_FORMAT(u.zeitpunkt,'%e.%c.%Y'),')') " & _
        "))) etext, " & _
        "d.icd, e.inhalt, (INSTR(pulsatp_re,'+')<>0 AND INSTR(pulsatp_li,'+')<>0) OR (INSTR(pulsadp_re,'+')<>0 AND INSTR(pulsadp_li,'+')<>0) unauff " & _
        "FROM namen n " & _
        "LEFT JOIN usdm u ON n.pat_id = u.pat_id AND u.zeitpunkt=(SELECT MAX(zeitpunkt) FROM usdm u1 WHERE u1.pat_id = u.pat_id) " & _
        "LEFT JOIN diagview d ON d.pat_id = n.pat_id AND d.gICD RLIKE '^I7[034]\.' AND d.id1 = (SELECT MAX(id1) FROM diagview d0 WHERE d0.pat_id = d.pat_id AND (d.gICD RLIKE '^I7[034]\.')) " & _
        "LEFT JOIN eintraege e ON e.pat_id = n.pat_id AND e.art IN ('doppler','duplex') AND e.inhalt LIKE '%beina%' AND e.zeitpunkt=(SELECT MAX(zeitpunkt) FROM eintraege e0 WHERE e0.pat_id = e.pat_id AND e.art IN ('doppler','duplex') AND e.inhalt LIKE '%beina%') " & _
        "WHERE n.Pat_id = " & Pat_ID
  myFrag rpuls, sql
  If Not rpuls.BOF Then
    PStatNeu = rpuls.Fields(1)
    etext = rpuls.Fields(2)
  End If
  Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
  AnwPfad = CurrentDb.name
#Else
  AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PStatNeu/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' PStatNeu

Function PStatus%(Pat_ID&, Optional frm As Object)
  Dim pul$, puk$, putp$, pudp$, pulp%, pukp%, putpp%, pudpp%, tprp As Boolean, tplp As Boolean, dprp As Boolean, dplp As Boolean
  On Error GoTo fehler
  If frm Is Nothing Then
'   SET frm = TabÖff("Anamnesebogen", "Pat_id")
'   frm.Seek "=", Pat_id
'   IF frm.NoMatch THEN Exit Function
   myFrag frm, "SELECT * FROM `anakt` WHERE pat_id = " & Pat_ID
   If frm.BOF Then Exit Function
  End If
  On Error Resume Next
  pul = IIf(IsNull(frm.anaRS![Puls Leiste]), vNS, LCase$(frm.anaRS![Puls Leiste]))
  puk = IIf(IsNull(frm.anaRS![Puls Kniekehle]), vNS, LCase$(frm.anaRS![Puls Kniekehle]))
  putp = IIf(IsNull(frm.anaRS![Puls Atp]), vNS, LCase$(frm.anaRS![Puls Atp]))
  pudp = IIf(IsNull(frm.anaRS![Puls Adp]), vNS, LCase$(frm.anaRS![Puls Adp]))
  On Error GoTo fehler
  If ((InStrB(pul, "-") <> 0 Or InStrB(pul, "mono") <> 0) And Not (putp = "+ | +" Or pudp = "+ | +" Or putp = "++ | ++" Or pudp = "++ | ++")) Or InStrB(pul, "pst") <> 0 Or InStrB(pul, "post") <> 0 Then
   Call doGilbE(frm, "I70.20", "vTextB", 121) ' Puls Leiste
   pulp = -1
   PStatus = -1
  End If
  If ((InStrB(puk, "-") <> 0 Or InStrB(puk, "mono") <> 0) And Not (putp = "+ | +" Or pudp = "+ | +" Or putp = "++ | ++" Or pudp = "++ | ++")) Or InStrB(puk, "pst") <> 0 Or InStrB(puk, "post") <> 0 Then
   Call doGilbE(frm, "I70.20", "vTextB", 122) ' Puls Kniekehle
   pukp = -1
   PStatus = -1
  End If
  Dim ttprp As Boolean, tdprp As Boolean, ttplp As Boolean, tdplp As Boolean
  Call PulsParse(putp, ttprp, ttplp)
  Call PulsParse(pudp, tdprp, tdplp)
  If (ttprp And tdprp) Or (ttplp And tdplp) Then
   PStatus = -1
   Call doGilbE(frm, "I70.20", "vTextB", 123) ' Puls Atp")
   Call doGilbE(frm, "I70.20", "vTextB", 124) ' Puls Adp")
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PStatus/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' PStatus
 
 Function obLH%(Pat_ID&, Optional frm As Object)
  Dim liph$
  On Error GoTo fehler
  If frm Is Nothing Then
'   SET frm = TabÖff("Anamnesebogen", "Pat_id")
'   frm.Seek "=", Pat_id
'   IF frm.NoMatch THEN Exit Function
   myFrag frm, "SELECT * FROM `anakt` WHERE pat_id = " & Pat_ID
   If frm.BOF Then Exit Function
  End If
  On Error Resume Next
  liph = IIf(IsNull(frm.anaRS![Liphypertrophien Abdomen]), vNS, LCase$(frm.anaRS![Liphypertrophien Abdomen]))
  obLH = Not obNein(liph) And Not obUnbek(liph) And liph <> "entfällt" And liph <> "kein insulin"
  On Error GoTo fehler
  If Not obLH Then
   liph = IIf(IsNull(frm.anaRS![Liphypertrophien Beine]), vNS, LCase$(frm.anaRS![Liphypertrophien Beine]))
   obLH = Not obNein(liph)
   If Not obLH Then
    liph = IIf(IsNull(frm.anaRS![Liphypertrophien Arme]), vNS, LCase$(frm.anaRS![Liphypertrophien Arme]))
    obLH = Not obNein(liph)
    If obLH Then
     Call doGilbE(frm, "L94.8", "vTextB", 108) 'Liphypertrophien Arme")
    End If
   Else
    Call doGilbE(frm, "L94.8", "vTextB", 107) 'Liphypertrophien Beine")
   End If
  Else
   Call doGilbE(frm, "L94.8", "vTextB", 106) 'Liphypertrophien Abdomen")
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in obLH/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' obLH
 
 Function LegNPFest$(Pat_ID&, Optional frm As Object)
  Dim obNeurop%, obVNeurop%, NPGrund$
  Dim rE$, li$
  On Error GoTo fehler
  If frm Is Nothing Then
'   SET frm = TabÖff("Anamnesebogen", "Pat_id")
'   frm.Seek "=", Pat_id
'   IF frm.NoMatch THEN Exit Function
   myFrag frm, "SELECT * FROM `anakt` WHERE pat_id = " & Pat_ID
   If frm.BOF Then Exit Function
  End If ' frm is Nothing
 
 obNeurop = False
 obVNeurop = False
 NPGrund = vNS
 If Not IsNull(frm.anaRS![Vibration IK]) Then
  Call bruchteile(frm.anaRS![Vibration IK], rE, li) ' `Vibration IK`
  If IsNumeric(rE) Then
   If Val(rE) < 5 Then
    If Val(rE) < 4 Then
     obNeurop = True
    Else
     obVNeurop = True
    End If
    Call doGilbE(frm, "G63.2", "vTextB", 119) 'Vibration IK")
   End If
  End If
  If IsNumeric(li) Then
   If Val(li) < 5 Then
    If Val(li) < 4 Then
     obNeurop = True
    Else
     obVNeurop = True
    End If
    Call doGilbE(frm, "G63.2", "vTextB", 119) 'Vibration IK")
   End If
  End If
 End If
 If Not IsNull(frm.anaRS![Vibration Großzehe]) Then
  Call bruchteile(frm.anaRS![Vibration Großzehe], rE, li) ' [Vibration Großzehe]
  If IsNumeric(rE) Then
   If Val(rE) < 5 Then
    If Val(rE) < 4 Then
     obNeurop = True
    Else
     obVNeurop = True
    End If
    Call doGilbE(frm, "G63.2", "vTextB", 120) 'Vibration Großzehe")
   End If
  End If
  If IsNumeric(li) Then
   If Val(li) < 5 Then
    If Val(li) < 4 Then
     obNeurop = True
    Else
     obVNeurop = True
    End If
    Call doGilbE(frm, "G63.2", "vTextB", 120) 'Vibration Großzehe")
   End If
  End If
 End If
 If obNeurop Or obVNeurop Then NPGrund = "Vibr"
 
 Dim ameis$
 ameis = IIf(IsNull(frm.anaRS!Ameisenlaufen), vNS, LCase$(frm.anaRS!Ameisenlaufen))
 If Left$(ameis, 1) = "j" Then
  obNeurop = True
  NPGrund = NPGrund + IIf(LenB(NPGrund) = 0, vNS, ",") + "Ameis"
  Call doGilbE(frm, "G63.2", "vTextB", 89) 'Ameisenlaufen")
 End If

 Dim kraft$, obkraft
 obkraft = 0
 If Not IsNull(frm.anaRS![Kraft Zehenheber]) Then
  kraft = frm.anaRS![Kraft Zehenheber]
 Else
  kraft = ""
 End If
 If kraft = "vermindert" Or InStrB(kraft, "schwach") <> 0 Or InStrB(kraft, "gering ") <> 0 Then
   obVNeurop = True
   obkraft = -1
   Call doGilbE(frm, "G63.2", "vTextB", 111) 'Kraft Zehenheber")
 End If
 If Not IsNull(frm.anaRS![Kraft Zehenbeuger]) Then
  kraft = frm.anaRS![Kraft Zehenbeuger]
 Else
  kraft = ""
 End If
 If kraft = "vermindert" Or InStrB(kraft, "schwach") <> 0 Or InStrB(kraft, "gering ") <> 0 Then
   obVNeurop = True
   obkraft = -1
   Call doGilbE(frm, "G63.2", "vTextB", 112) 'Kraft Zehenbeuger")
 End If
 If Not IsNull(frm.anaRS![Kraft Knie]) Then
  kraft = frm.anaRS![Kraft Knie]
 Else
  kraft = ""
 End If
 If kraft = "vermindert" Or InStrB(kraft, "schwach") <> 0 Or InStrB(kraft, "gering ") <> 0 Then
   obVNeurop = True
   obkraft = -1
   Call doGilbE(frm, "G63.2", "vTextB", 113) 'Kraft Knie")
 End If
 If obkraft Then NPGrund = NPGrund + IIf(LenB(NPGrund) = 0, vNS, ",") + "Kra"
 
 Dim mer$, obmer As Boolean
 obmer = 0
 If Not IsNull(frm.anaRS!ASR) Then
  mer = frm.anaRS!ASR
 Else
  mer = ""
 End If
 If (InStrB(mer, "-") <> 0 And InStrB(mer, "?") <> 0) Or mer = "-" Or mer = "-/-" Then
  obVNeurop = True
  obmer = True
  Call doGilbE(frm, "G63.2", "vTextB", 114) 'ASR")
 End If
 If Not IsNull(frm.anaRS!PSR) Then
  mer = frm.anaRS!PSR
 Else
  mer = ""
 End If
 If (InStrB(mer, "-") <> 0 And InStrB(mer, "?") <> 0) Or mer = "-" Or mer = "-/-" Then
  obVNeurop = True
  obmer = True
  Call doGilbE(frm, "G63.2", "vTextB", 115) 'PSR")
 End If
 If obmer Then NPGrund = NPGrund + IIf(LenB(NPGrund) = 0, vNS, ",") + "MER"
  
 Dim ofl$, obofl As Boolean
 If Not IsNull(frm.anaRS!Oberflächensensibilität) Then
  ofl = frm.anaRS!Oberflächensensibilität
 Else
  ofl = ""
 End If
 If InStrB(ofl, "gestört") <> 0 Or InStrB(ofl, "paret") <> 0 Or InStrB(ofl, "vermind") <> 0 Or InStrB(ofl, "path") <> 0 Or InStrB(ofl, "pelz") <> 0 Or InStrB(ofl, "strumpf") <> 0 Or InStrB(ofl, "schwäch") <> 0 Or InStrB(ofl, "eingeschr") <> 0 Then
  obNeurop = True
  NPGrund = NPGrund + IIf(LenB(NPGrund) = 0, vNS, ",") + "Ofl"
  Call doGilbE(frm, "G63.2", "vTextB", 116) 'Oberflächensensibilität")
 End If
 
 Dim Monf$
 If Not IsNull(frm.anaRS!Monofilamenttest) Then
  Monf = frm.anaRS!Monofilamenttest
 Else
  Monf = ""
 End If
 If Left$(Monf, 1) = "0" Or InStrB(Monf, " 0") <> 0 Or InStrB(Monf, "1") <> 0 Or InStrB(Monf, "2") <> 0 Or (InStrB(Monf, "3") <> 0 And InStrB(Monf, "/3") = 0) Or InStrB(Monf, "4") <> 0 Or InStrB(Monf, "6/8") <> 0 Or InStrB(Monf, "5/8") <> 0 Or InStrB(Monf, "gestört") <> 0 Or InStrB(Monf, "5/10") <> 0 Or InStrB(Monf, "6/10") <> 0 Or InStrB(Monf, "7/10") <> 0 Or InStrB(Monf, "8/10") <> 0 Or InStrB(Monf, "neg") <> 0 Or InStrB(Monf, "path") <> 0 Or InStrB(Monf, " -") <> 0 Or InStrB(Monf, "zufäl") <> 0 Or InStrB(Monf, "erat") <> 0 Then
  obNeurop = True
  NPGrund = NPGrund + IIf(LenB(NPGrund) = 0, vNS, ",") + "MF"
  Call doGilbE(frm, "G63.2", "vTextB", 117) 'Monofilamenttest")
 End If
 
 Dim KW$
 If Not IsNull(frm.anaRS![Kalt-Warm]) Then
  KW = frm.anaRS![Kalt-Warm]
 Else
  KW = ""
 End If
 If Left$(KW, 1) = "0" Or InStrB(KW, " 0") <> 0 Or InStrB(KW, "1") <> 0 Or InStrB(KW, "2") <> 0 Then
' OR instrb(Monf, "6/8") <> 0 OR instrb(Monf, "5/8") <> 0 OR instrb(Monf, "6/8") <> 0 OR instrb(Monf, "5/10") <> 0 OR instrb(Monf, "6/10") <> 0 OR instrb(Monf, "7/10") <> 0 OR instrb(KW, "gestört") <> 0 OR instrb(KW, "neg") <> 0 OR instrb(KW, "path") <> 0 OR instrb(KW, " -") <> 0 OR instrb(KW, "zufäl") <> 0 OR instrb(KW, "erat") <> 0 THEN ' OR instrb(KW, "4") <> 0 OR instrb(Monf, "8/10") <> 0
  obVNeurop = True
  NPGrund = NPGrund + IIf(LenB(NPGrund) = 0, vNS, ",") + "KW"
  Call doGilbE(frm, "G63.2", "vTextB", 118) 'Kalt-Warm")
 End If
 LegNPFest = NPGrund
' damit nur ein Parameter zugegeben zu werden braucht
 If obVNeurop Then
  LegNPFest = "V." + LegNPFest
 ElseIf obNeurop Then
  LegNPFest = "->" + LegNPFest
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LegNPFest/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LegNPFest
#End If

' wird sowohl von form_current als auch von einer dynamischen SQL-Abfrage aufgerufen, deshalb zwei Aufrufmöglichkeiten
 Function hkGrund$(Pat_ID&, Optional rsNa As Object)
  Dim druckst$, hyperkerat$, obDst%, obHk%, i%, hkspl$()
  On Error GoTo fehler
   If rsNa Is Nothing Then
'    SET rsNa = TabÖff("Anamnesebogen", "Pat_id")
'    rsNa.Seek "=", Pat_id
'    IF rsNa.NoMatch THEN Exit Function
    Set rsNa = New ADODB.Recordset
    myFrag rsNa, "SELECT * FROM `anamnesebogen` WHERE pat_id = " & Pat_ID, adOpenStatic
    If rsNa.BOF Then Exit Function
   End If
'  IF Not rsNa.NoMatch THEN
   obDst = 0
   obHk = 0
   hkGrund = vNS
   druckst = IIf(IsNull(rsNa!Druckstellen), vNS, LCase$(rsNa!Druckstellen))
   hyperkerat = IIf(IsNull(rsNa!Hyperkeratosen), vNS, LCase$(rsNa!Hyperkeratosen))
   hyperkerat = Trim$(hyperkerat)
   If Left$(druckst, 1) = "j" Then
    obDst = -1
    hkGrund = "Frgbg."
   End If
   If obPosi(hyperkerat) Then
    If InStrB(hyperkerat, "|") <> 0 Then
     hkspl = Split(hyperkerat, "|")
     For i = 0 To UBound(hkspl)
      If InStrB(hkspl(i), "leicht") = 0 Then
       If InStrB(hkspl(i), "etwas") = 0 Then
        If obPosi(hkspl(i)) Then
         obHk = -1
         Exit For
        End If ' obPosi(hkspl(i)) THEN
       End If ' InStrB(hkspl(i), "etwas") = 0 THEN
      End If ' InStrB(hkspl(i), "leicht") = 0 THEN
     Next i ' i = 0 To UBound(hkspl)
    End If ' IF InStrB(hyperkerat, "|") <> 0 THEN
    If obHk Then
     hkGrund = IIf(LenB(hkGrund) = 0, vNS, hkGrund + ",") + "Befd"
    End If
   End If
 ' END IF
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in hkGrund/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' hkGrund
 
#If False Then
' wird von SQL-Abfragen aufgerufen
  Function EigTAlter$(Pat_ID&, Eig$, vz$, grenze$)
   Dim Alter$, GebDat#, obErfüllt%, rsNa As New ADODB.Recordset
   On Error GoTo fehler
'   SET rNa = TabÖff("Namen", "Pat_id")
'   rNa.Seek "=", Pat_id
   myFrag rsNa, "SELECT * FROM `namen` WHERE pat_id = " & Pat_ID
   If Not rsNa.BOF Then
    GebDat = rsNa!GebDat
    Alter = (Date - GebDat) * 0.0027379
    Select Case vz
     Case ">"
      If Alter > grenze Then obErfüllt = -1
     Case "<"
      If Alter < grenze Then obErfüllt = -1
    End Select
    If obErfüllt Then
     EigTAlter = EigT(Pat_ID, Eig)
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in EigTAlter/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' EigTAlter$(Pat_id&, Eig$, vz$, grenze$)
#End If

#If False Then
' wird von SQL-Abfragen aufgerufen
  Function EigT$(Pat_ID&, Eig$) ' Eigenschaft testen
   Static sql$, altpatE&
   Static md As DAO.Recordset
   On Error GoTo fehler
   Call dtbInit
   If Pat_ID = 0 Then Exit Function
   If Pat_ID <> altpatE Then
    sql = "SELECT * FROM `" + QMdbAkt + "`.`medikamente mit Arten` WHERE `medplan`.pat_id =" + CStr(Pat_ID)
    Set md = Dtb.OpenRecordset(sql, dbOpenDynaset)
    altpatE = Pat_ID
   End If
   If Not md.BOF Then
    md.MoveFirst
    Do While Not md.EOF
     Call mdTest(md, Eig, EigT)
     md.Move 1
    Loop
   End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in EigT/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
  End Function
#End If

' in do_An1Aufruf_click, do_An2Aufruf_click, do_AnAAufruf_click, do_An1Aufruf_click, do_CheckAufruf_click, do_Form_Current2
 Function PfadFestLeg$(art$, muster1$, frm As Form, Optional muster2$)
  Dim PFrs As ADODB.Recordset
  On Error GoTo fehler
  sql = "SELECT * FROM briefe WHERE pat_id = " + CStr(frm.anaRS!Pat_ID) + " AND (name LIKE " & "'" & muster1 & "'" & IIf(LenB(muster2) = 0, vbNullChar, " OR name LIKE " & "'" & muster2 + "'") + ") ORDER BY zeitpunkt DESC"
  On Error GoTo fehler
'  Call dtbInit
'  SET rDok = Dtb.OpenRecordset(sql, dbOpenDynaset)
  Set PFrs = Nothing
  Exit Function
  myFrag PFrs, sql
  If PcDokPfad = vNS Then Call getDokPfad
  Dim ctl, fehler&
  For Each ctl In frm.Controls
   Dim dataf$
   On Error Resume Next
   Err.Clear
   dataf = ctl.DataFields
   fehler = Err.Number
   On Error GoTo fehler
   If fehler = 0 Then
    If ctl.DataField Like art Then
     If PFrs.BOF Or PcDokPfad = vNS Then
      frm.Controls(ctl).Enabled = False
     Else
      frm.Controls(ctl).Enabled = True
  '   PFrs.MoveFirst
      PfadFestLeg = PFrs!DokPfad
      PFrs.Close
     End If
    End If
    Exit For
   End If
  Next ctl
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PfadFestLeg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' PfadFestLeg$(Art$, muster1$, frm As Form, Optional muster2$)

Function Urineintraege%(Pat_ID&)
 Dim pZahlStix%, pZahlLab%, nZahlStix%, nZahlLab%, nZahlMic%, pZahlMic%, Ps$(), i%
 Dim raUr As ADODB.Recordset
 On Error GoTo fehler
 Set raUr = New ADODB.Recordset
' Call raUr.Open("SELECT * FROM `eintraege` WHERE art = ""urin"" AND pat_id = " & CStr(Pat_id), DBCn, adOpenDynamic, adLockReadOnly)
 myFrag raUr, "SELECT * FROM `eintraege` WHERE art = ""urin"" AND pat_id = " & CStr(Pat_ID)
 Do While Not raUr.EOF
  Ps = Split(raUr!Inhalt, ",")
   For i = 0 To UBound(Ps())
    If InStrB(Ps(i), "Eiwei") <> 0 Then
     If InStrB(Ps(i), "+") <> 0 Then
      pZahlStix = pZahlStix + 1
     ElseIf InStrB(Ps(i), "-") <> 0 Then
      nZahlStix = nZahlStix + 1
     End If
    ElseIf InStrB(Ps(i), "Micral") <> 0 Then
     If (InStrB(Ps(i), "20") <> 0 Or InStrB(Ps(i), "50") <> 0) Then
      pZahlMic = pZahlMic + 1
     ElseIf InStrB(Ps(i), "neg") <> 0 Then
      nZahlMic = nZahlMic + 1
     End If
    End If
   Next i
  raUr.MoveNext
 Loop
 Dim LW&
' rlau.Seek "=", CStr(Pat_id), "albcre"
' raUr.Close
' SET raUr = New ADODB.Recordset
' Call raUr.Open("SELECT * FROM `laborneu` LEFT JOIN laborkommentar ON `laborneu`.kommentarvw = laborkommentar.kommentarvw WHERE abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBU') AND (abkü <> 'ALBU' OR wert LIKE '%<%') AND pat_id = " & CStr(Pat_ID), DBCn, adOpenDynamic, adLockReadOnly)
' SET raUr = LabEPat(AlbCre, Pat_id)
 Dim Labs As labtyp
 alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
 Do While True
  Labs = LabPat(LA_AlbCre, Pat_ID, True)
  If Labs.Zp = -1 Or Labs.Zp = #12/29/1899# Then Exit Do
' IF Not raUr.BOF THEN
'  Do While Not raUr.EOF
''   LW = Val(replace$(replace$(IIF(ISNULL(raUr!Wert), IIF(ISNULL(raUr!Kommentar), vns, raUr!Kommentar), raUr!Wert), ",", "."), "%", vns))
  LW = Val(REPLACE$(REPLACE$(Labs.WertSg, ",", "."), "%", vNS))
  If LW < 20 Then
   nZahlLab = nZahlLab + 1
  Else
   pZahlLab = pZahlLab + 1
  End If
'   raUr.Move 1
 Loop
 alt_la = LabArt0 ' damit beim nä gleichartigen Aufruf nicht nichts rauskommt
 Urineintraege = pZahlStix + pZahlMic - nZahlMic - nZahlMic - nZahlLab - nZahlLab - nZahlLab ' nzahlmic zählt doppelt, nzahllab dreifach
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Urineintraege/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Urineintraege

Function doXtra_Click(frm As AnBog)
 Dim rAf&
 On Error GoTo fehler
 If UStumm Then Exit Function
 If frm.Xtra <> vNS Then
 Call myEFrag("UPDATE `fuerdiagexp` SET icd = CONCAT(LEFT(icd,instr(icd,'.')-1),'" & Left$(frm.Xtra, 3) & "'), zeitpunkt = " & DatFor_k(Now()) & " WHERE pat_id = " & frm.anaRS!Pat_ID & " ORDER BY id DESC LIMIT 1", rAf)  ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call myEFrag("UPDATE `fuerdiagexp` SET diagnose = MID(diagnose,6), zeitpunkt = " & DatFor_k(Now()) & " WHERE pat_id = " & frm.anaRS!Pat_ID & " AND (diagnose LIKE 'Z.n.%' OR diagnose LIKE 'Z.n.%V.a.%')  ORDER BY id DESC LIMIT 1", rAf) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call myEFrag("UPDATE `fuerdiagexp` SET diagnose = MID(diagnose,6), zeitpunkt = " & DatFor_k(Now()) & " WHERE pat_id = " & frm.anaRS!Pat_ID & " AND (diagnose LIKE 'V.a.%' OR diagnose LIKE 'V.a.%Z.n.%')  ORDER BY id DESC LIMIT 1", rAf) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call myEFrag("UPDATE `fuerdiagexp` SET diagnose = MID(diagnose,6), zeitpunkt = " & DatFor_k(Now()) & " WHERE pat_id = " & frm.anaRS!Pat_ID & " AND (diagnose LIKE 'Z.n.%' OR diagnose LIKE 'Z.n.%V.a.%')  ORDER BY id DESC LIMIT 1", rAf) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call myEFrag("UPDATE `fuerdiagexp` SET diagnose = MID(diagnose,6), zeitpunkt = " & DatFor_k(Now()) & " WHERE pat_id = " & frm.anaRS!Pat_ID & " AND (diagnose LIKE 'V.a.%' OR diagnose LIKE 'V.a.%Z.n.%')  ORDER BY id DESC LIMIT 1", rAf) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  UStumm = True
  frm.Va = 0
  frm.Zn = 0
  UStumm = False
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doVa_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doXtra_Click

Function doVa_Click(frm As Form)
 Dim rAf&
 On Error GoTo fehler
 If UStumm Then Exit Function
 If frm.Va Then
  Call myEFrag("UPDATE `fuerdiagexp` SET icd = CONCAT(icd,'V'), zeitpunkt = " & DatFor_k(Now()) & " WHERE pat_id = " & frm.anaRS!Pat_ID & " AND (not icd LIKE '%V' AND NOT icd LIKE '%VZ')  ORDER BY id DESC LIMIT 1", rAf) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call myEFrag("UPDATE `fuerdiagexp` SET diagnose = CONCAT('V.a. ',diagnose), zeitpunkt = " & DatFor_k(Now()) & " WHERE pat_id = " & frm.anaRS!Pat_ID & " AND (not diagnose LIKE 'V.a.%' AND NOT diagnose LIKE 'V.a.%Z.n.%')  ORDER BY id DESC LIMIT 1", rAf) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
 Else
  Call myEFrag("UPDATE `fuerdiagexp` SET icd = LEFT(icd,LENGTH(icd)-1), zeitpunkt = " & DatFor_k(Now()) & " WHERE pat_id = " & frm.anaRS!Pat_ID & " AND (icd LIKE '%V' OR icd LIKE '%VZ') ORDER BY id DESC LIMIT 1", rAf) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call myEFrag("UPDATE `fuerdiagexp` SET diagnose = ltrim(MID(diagnose,5)), zeitpunkt = " & DatFor_k(Now()) & " WHERE pat_id = " & frm.anaRS!Pat_ID & " AND (diagnose LIKE 'V.a.%' OR diagnose LIKE 'V.a.%Z.n.%') ORDER BY id DESC LIMIT 1", rAf) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doVa_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doVa_Click

Function doZn_Click(frm As Form)
 Dim rAf&
 On Error GoTo fehler
 If UStumm Then Exit Function
 If frm.Zn Then
  Call myEFrag("UPDATE `fuerdiagexp` SET icd = CONCAT(icd,'Z'), zeitpunkt = " & DatFor_k(Now()) & " WHERE pat_id = " & frm.anaRS!Pat_ID & " AND (not icd LIKE '%Z' AND NOT icd LIKE '%ZV') ORDER BY id DESC LIMIT 1", rAf) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call myEFrag("UPDATE `fuerdiagexp` SET diagnose = CONCAT('Z.n. ',diagnose), zeitpunkt = " & DatFor_k(Now()) & " WHERE pat_id = " & frm.anaRS!Pat_ID & " AND (not diagnose LIKE 'Z.n.%' AND NOT diagnose LIKE 'Z.n.%V.a.%') ORDER BY id DESC LIMIT 1", rAf) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
 Else
  Call myEFrag("UPDATE `fuerdiagexp` SET icd = LEFT(icd,LENGTH(icd)-1), zeitpunkt = " & DatFor_k(Now()) & " WHERE pat_id = " & frm.anaRS!Pat_ID & " AND (icd LIKE '%Z' OR icd LIKE '%ZV') ORDER BY id DESC LIMIT 1", rAf) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
  Call myEFrag("UPDATE `fuerdiagexp` SET diagnose = ltrim(MID(diagnose,5)) WHERE pat_id = " & frm.anaRS!Pat_ID & " AND (diagnose LIKE 'Z.n.%' OR diagnose LIKE 'Z.n.%V.a.%')  ORDER BY id DESC LIMIT 1", rAf) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doZn_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doZn_Click

Function knöpfeanpassen(frm As AnBog)
 Dim nr%, i%
 Dim rfDE As New ADODB.Recordset
 On Error GoTo fehler
 UStumm = True
 frm.Va = 0
 frm.Zn = 0
 frm.Xtra = vNS
 For nr = 1 To 23
' Hier noch einfügen, wenn Tabelle rfDE abgefragt werden soll
  Dim icdlike$
  If MDIICD(nr) = vNS Then
   If frm.Controls("vOptionB")(nr).Value <> 0 Then frm.Controls("vOptionB")(nr).Value = False      ' "u"
  Else
   If MDIICD(nr) Like "E1*" Or MDIICD(nr) Like "L89*" Then
    icdlike = Left$(MDIICD(nr), InStr(MDIICD(nr), ".")) & ".*"
   Else
    icdlike = MDIICD(nr)
    For i = 1 To 2
     If Right$(icdlike, 1) = "V" Then icdlike = Left$(icdlike, Len(icdlike) - 1)
     If Right$(icdlike, 1) = "Z" Then icdlike = Left$(icdlike, Len(icdlike) - 1)
    Next i
    icdlike = icdlike & "[VZ]?"
   End If
   Set rfDE = Nothing
   If LVobMySQL Then
'    Call rfDE.Open("SELECT * FROM `fuerdiagexp` WHERE ISNULL(status) AND pat_id = " & frm.anaRS!Pat_id & " AND icd RLIKE '" & icdlike & "' AND diagnose LIKE '%" & LTrim$(replace$(LTrim$(replace$(MDIDiag(nr), "V.a.", vNS)), "Z.n.", vNS)) & "'", DBCn, adOpenDynamic, adLockReadOnly) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
    myFrag rfDE, "SELECT * FROM `fuerdiagexp` WHERE ISNULL(status) AND pat_id = " & frm.anaRS!Pat_ID & " AND icd RLIKE '" & icdlike & "' AND diagnose LIKE '%" & LTrim$(REPLACE$(LTrim$(REPLACE$(MDIDiag(nr), "V.a.", vNS)), "Z.n.", vNS)) & "'"
   Else
'    Call rfDE.Open("SELECT * FROM `fuerdiagexp` WHERE ISNULL(status) AND pat_id = " & frm.anaRS!Pat_id & " AND (icd = '" & MDIICD(nr) & "' OR icd = '" & MDIICD(nr) & "V' OR icd = '" & MDIICD(nr) & "Z' OR icd = '" & MDIICD(nr) & "VZ' OR icd = '" & MDIICD(nr) & "ZV') AND diagnose LIKE '%" & LTrim$(replace$(LTrim$(replace$(MDIDiag(nr), "V.a.", vNS)), "Z.n.", vNS)) & "'", DBCn, adOpenDynamic, adLockReadOnly) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
    myFrag rfDE, "SELECT * FROM `fuerdiagexp` WHERE ISNULL(status) AND pat_id = " & frm.anaRS!Pat_ID & " AND (icd = '" & MDIICD(nr) & "' OR icd = '" & MDIICD(nr) & "V' OR icd = '" & MDIICD(nr) & "Z' OR icd = '" & MDIICD(nr) & "VZ' OR icd = '" & MDIICD(nr) & "ZV') AND diagnose LIKE '%" & LTrim$(REPLACE$(LTrim$(REPLACE$(MDIDiag(nr), "V.a.", vNS)), "Z.n.", vNS)) & "'"
   End If
'  rfDE.Seek "=", frm.anaRS!Pat_id, MDIICD(nr), MDIDiag(nr)
   If (rfDE.BOF Or Trim$(MDIICD(nr)) = vNS) Then
    If frm.Controls("vOptionB")(nr).Value <> 0 Then frm.Controls("vOptionB")(nr).Value = 0  ' "u"
   Else
    If frm.Controls("vOptionB")(nr).Value = 0 Then frm.Controls("vOptionB")(nr).Value = 1 ' "u"
   End If
  End If
 Next nr
 Dim diabzahl&, passnr%
 diabzahl = 0
 passnr = 0
 For nr = 1 To 23
  If MDIICD(nr) Like "E1*" Then
   diabzahl = diabzahl + 1
  End If
 Next nr
 If diabzahl > 1 Then
  For nr = 1 To 23
   If MDIICD(nr) Like "E1*" Then
    icdlike = Left$(MDIICD(nr), 6) & "%"
    Set rfDE = Nothing
'    Call rfDE.Open("SELECT * FROM `fuerdiagexp` WHERE pat_id = " & frm.anaRS!Pat_id & " AND icd LIKE '" & icdlike & "'", DBCn, adOpenDynamic, adLockReadOnly) ' rfDE = TabÖff("fuerDiagExp", "Suche") ' pat_id, icd
    myFrag rfDE, "SELECT * FROM `fuerdiagexp` WHERE pat_id = " & frm.anaRS!Pat_ID & " AND icd LIKE '" & icdlike & "'"
    If Not rfDE.EOF Then
     passnr = nr
     Exit For
    End If
   End If
  Next nr
  If passnr = 0 Then
   For nr = 1 To 23
    If MDIICD(nr) Like "E1*" Then
     diabzahl = diabzahl - 1
     frm.Controls("vOptionB")(nr).Value = 0
     If diabzahl = 1 Then Exit For
    End If
   Next nr
  Else
   For nr = 1 To 23
    If MDIICD(nr) Like "E1*" Then
     If nr <> passnr Then
      diabzahl = diabzahl - 1
      frm.Controls("vOptionB")(nr).Value = 0
     End If
    End If
   Next nr
  End If
 End If ' diabzahl > 0
 UStumm = 0
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in knöpfeanpassen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' knöpfeanpassen

#If zutesten Then
Function RRpAll()
Dim Pat_ID&, rsNam As New ADODB.Recordset
On Error GoTo fehler
Call Lese.ProgStart
myFrag rsNam, "SELECT pat_id FROM `namen` ORDER BY pat_id"
Do While Not rsNam.EOF
 Call RRParseF(rsNam!Pat_ID)
 rsNam.Move 1
Loop
Call Lese.ProgEnde
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in RRpAll/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'RRpAll
#End If

' aufgerufen in do_Form_AfterUpdate, RRpAll
Function RRParseF(Pat_ID&)
 Dim rsRR As New ADODB.Recordset, rsP As New ADODB.Recordset, rsNa As New ADODB.Recordset
 On Error GoTo fehler
 Call myEFrag("DELETE FROM rrparse WHERE pat_id = " & Pat_ID)
' rsRR.Open "SELECT * FROM rr WHERE pat_id = " & CStr(Pat_id), DBCn, adOpenDynamic, adLockReadOnly
 myFrag rsRR, "SELECT RR,Pat_id,Zeitpunkt,""Tabelle RR"" FROM rr WHERE pat_id = " & CStr(Pat_ID)
 Do While Not rsRR.EOF
  Call do_RRParse(rsRR!RR, rsRR!Pat_ID, rsRR!Zeitpunkt, "Tabelle RR")
  rsRR.MoveNext
 Loop
 Set rsNa = Nothing
' rsNa.Open "SELECT * FROM `anamnesebogen` WHERE pat_id = " & Pat_id, DBCn, adOpenDynamic, adLockOptimistic
 myFrag rsNa, "SELECT * FROM `anamnesebogen` WHERE pat_id = " & Pat_ID
 If Not rsNa.EOF Then
  If Not IsNull(rsNa!Blutdruckwerte) Then
   Call do_RRParse(rsNa!Blutdruckwerte, Pat_ID, IIf(IsNull(rsNa!Vorgestellt), 0, rsNa!Vorgestellt), "An'bg.BW")
  End If
  If Not IsNull(rsNa!RR) Then
   Call do_RRParse(rsNa!RR, Pat_ID, rsNa!Vorgestellt, "An'bg.RR")
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in RRParse/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' RRParse()

' 20.3.17: folgende Funktion könnte jetzt durch SQL-Funktionen RRsyst und RRdiast erübrigt werden
' in RRParse und rrParseSpeichern
Function do_RRParse(erg$, ByVal Pat_ID, Zeitpunkt As Date, Quelle$)
  Dim RRsyst%, RRdiast%, Zp As Date, rAf&
  Dim rsP As New ADODB.Recordset
  On Error GoTo fehler
  If dodoRRParse(erg, RRsyst, RRdiast, Zp) = 0 Then Exit Function
'  rsP.Seek "=", Pat_id, IIF(Zp = 0, Zeitpunkt, Zp), RRSyst, RRDiast
  myFrag rsP, "SELECT 0 FROM `rrparse` WHERE pat_id = " & Pat_ID & " AND zeitpunkt = " & DatFor_k(IIf(Zp = 0, Zeitpunkt, Zp)) & " AND rrsyst = " & RRsyst & " AND rrdiast = " & RRdiast
  If rsP.BOF Then
   InsKorr DBCn, "INSERT INTO `rrparse`(pat_id,zeitpunkt,rrsyst,rrdiast,quelle) VALUES(" & Pat_ID & "," & DatFor_k(IIf(Zp = 0, Zeitpunkt, Zp)) & "," & RRsyst & "," & RRdiast & ",'" & REPLACE$(Quelle, "'", "''") & "')", rAf
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_RRParse/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_RRParse(erg$, rsP AS dao.recordset, Pat_id$, ZeitPunkt As Date)

#If zutesten Then
Function testrp(p$)
 Dim s%, d%
 Call dodoRRParse(p, s, d)
 Debug.Print s, d
End Function ' testrp(p$)
#End If

' 20.3.17: folgende Funktion sollte jetzt durch SQL-Funktionen RRsyst und RRdiast erübrigt sein
' aufgerufen in rrParsen, und GetPrRR, do_RRParse
Function dodoRRParse(ByVal erg$, RRsyst%, RRdiast%, Optional Zp As Date, Optional Puls%, Optional Bem$)
  Dim i%, RR1$, RR1a$, pos%, runde%
  Dim pko%
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
    RR1 = Left$(erg, pos - 1)
    If Len(RR1) > 2 Then
'     IF Mid$(RR1, Len(RR1) - 2, 1) = "," AND InStrB("0123456789", Mid$(RR1, Len(RR1) - 1, 1)) <> 0 THEN RR1 = LEFT(RR1, Len(RR1) - 2)
     pko = InStr(RR1, ",")
     If pko > 1 And IsNumeric(Mid$(RR1, pko + 1, 1)) Then RR1 = Left$(RR1, pko - 1)
    End If
    RR1a = Right$(RR1, 3)
    If IsNumeric(RR1a) Then
     RRsyst = Val(RR1a)
    Else
     RR1a = Right$(RR1a, 2)
     If IsNumeric(RR1a) Then
      RRsyst = Val(RR1a)
     End If
    End If
    erg = Mid$(erg, pos + 1)
    RR1a = Left$(erg, 3)
    If IsNumeric(RR1a) Then
     RRdiast = Val(RR1a)
     erg = Mid$(erg, 4)
    Else
     RR1a = Left$(RR1a, 2)
     If IsNumeric(RR1a) Then
      RRdiast = Val(RR1a)
      erg = Mid$(erg, 2)
     End If
    End If
    If InStr(erg, "(") > InStr(erg, "/") Then
     RR1a = Mid$(erg, InStr(erg, "("))
     If InStr(RR1a, ")") > 1 Then
      RR1a = Left$(RR1a, InStr(RR1a, ")") - 1)
      If IsDate(RR1a) Then Zp = CDate(RR1a)
     End If
    End If
   Else ' pos > 1
    pos = InStr(erg, "syst")
    If pos > 1 Then
     RR1 = Left$(erg, pos - 1)
     For i = Len(RR1) To 0 Step -1
      If InStrB("0123456789", Mid$(RR1, i, 1)) <> 0 And LenB(Mid$(RR1, i, 1)) <> 0 Then Exit For
      RR1 = Left$(RR1, Len(RR1) - 1)
      If RR1 = vNS Then Exit For
     Next
     If Len(RR1) > 2 Then
      If Mid$(RR1, Len(RR1) - 2, 1) = "," And InStrB("0123456789", Mid$(RR1, Len(RR1) - 1, 1)) <> 0 And Mid$(RR1, Len(RR1) - 1, 1) <> vNS Then RR1 = Left$(RR1, Len(RR1) - 2)
     End If
     If RR1 = vNS Then GoTo nix
     RR1a = Right$(RR1, 3)
     If IsNumeric(RR1a) Then
      RRsyst = Val(RR1a)
     Else
      RR1a = Right$(RR1a, 2)
      If IsNumeric(RR1a) Then
       RRsyst = Val(RR1a)
      End If
      If RRsyst < 50 Then
       RR1a = Left$(RR1a, Len(RR1a) - 2)
       For i = Len(RR1a) To 1 Step -1
        If InStrB("0123456789", Mid$(RR1a, i, 1)) <> 0 And Mid$(RR1a, i, 1) <> vNS Then Exit For
        RR1 = Left$(RR1a, Len(RR1a) - 1)
       Next
       If Len(RR1a) > 2 Then
        If Mid$(RR1a, Len(RR1a) - 2, 1) = "," And InStrB("0123456789", Mid$(RR1a, Len(RR1a) - 1, 1)) <> 0 And Mid$(RR1a, Len(RR1a) - 1, 1) <> vNS Then RR1a = Left$(RR1a, Len(RR1a) - 2)
       End If
       RR1a = Right$(RR1a, 3)
       If IsNumeric(RR1a) Then
        RRsyst = Val(RR1a)
       End If
      End If
     End If
     erg = Mid$(erg, pos + 4)
    Else ' pos > 1
     erg = vNS
    End If ' pos > 1
   End If ' pos > 1
   If RRsyst > 0 Then Exit Do
   If erg = vNS Then
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

Function mdTest(md As ADODB.Recordset, Feld$, Optional ByRef bisher$)
 Dim medi$
 On Error GoTo fehler
 If Not IsNull(md.Fields(Feld).Value) Then
  If md.Fields(Feld).Value Then
   If LVobMySQL Then
    medi = Left$(md!Medikament, 15)
   Else
    medi = Left$(md![mp.Medikament], 15)
   End If
   If IsNull(bisher) Or bisher = vNS Then
     bisher = medi
   Else
    If Not InStrB(bisher, medi) <> 0 Then
     bisher = bisher & ", " & medi
    End If
   End If
  End If
 End If
 mdTest = bisher
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in mdTEst/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' mdTest

#If False Then
Function mdTest_alt$(md As DAO.Recordset, Feld$)
 On Error GoTo fehler
 If Not IsNull(md.Fields(Feld$).Value) Then
  If md.Fields(Feld$).Value Then
   If IsNull(mdTest_alt) Or mdTest_alt = vNS Then
     mdTest_alt = md!medi
   Else
     mdTest_alt = mdTest_alt & ", " & md!medi
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in mdTest_alt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' mdTest

#End If
Function do_An1Aufruf_click(frm As Form)
 On Error GoTo fehler
 If An1Pfad = vNS Then An1Pfad = PfadFestLeg("An1Aufruf", "%anamnese%1%", frm)
 Call do_do_Aufruf(frm, An1Pfad)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_An1Aufruf_click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_An1Aufruf_click(frm AS Form)

Function do_An2Aufruf_click(frm As Form)
 On Error GoTo fehler
 If An2Pfad = vNS Then An2Pfad = PfadFestLeg("An2Aufruf", "%anamnese%2%", frm)
 Call do_do_Aufruf(frm, An2Pfad)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_An2Aufruf_click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_An2Aufruf_click(frm As Form)

Function do_AnAAufruf_click(frm As Form)
 On Error GoTo fehler
 If AnAPfad = vNS Then AnAPfad = PfadFestLeg("AnAAufruf", "%anamnese%allg%", frm, "*allg*anamnese*")
 Call do_do_Aufruf(frm, AnAPfad)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_AnAAufruf_click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_AnAAufruf_click(frm As Form)

Function do_CheckAufruf_click(frm As Form)
 On Error GoTo fehler
 If CheckPfad = vNS Then CheckPfad = PfadFestLeg("CheckAufruf", "%che%kliste%", frm)
 Call do_do_Aufruf(frm, CheckPfad)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_CheckAufruf_click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_CheckAufruf_click(frm As Form)

Function do_do_Aufruf(frm As Form, DPfad$)
 On Error GoTo fehler
 If (InStrB("$\", Left$(DPfad, 1)) <> 0 And Left$(DPfad, 1) <> vNS) Or Mid$(DPfad, 2, 1) = ":" Then
   If LenB(PcDokPfad) = 0 Then getDokPfad
   DatPfad = getIViewPfad & " " & vNS & Stutz(REPLACE$(LCase$(DPfad), "$\turbomed\dokumente", PcDokPfad)) + vNS ' Environ("ProgramFiles") & "\irfanview\i_view32.exe
   Call Shell(DatPfad, vbMaximizedFocus)
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_do_Aufruf/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_do_Aufruf(frm As Form, DPfad$)

Function inActi(ct As Control)
  On Error Resume Next
  ct.BackColor = 12632256 ' grau
  On Error GoTo fehler
  ct.Enabled = False
  On Error Resume Next
  ct.Locked = True
  On Error GoTo fehler
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in inActi/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function

Function Acti(ct As Control)
  On Error Resume Next
  ct.BackColor = 16777215 ' weiß
  If ct.name <> "vCheckb" Then ct.Locked = False ' wg. häufigen Hängenbleibens beim Debugenn mit Err-Überwachung
  On Error GoTo fehler
  ct.Enabled = True
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Acti/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Acti(ct As Control)

Function do_Hirndurchblutungsstörung_Exit(Cancel%, frm As Form)
 On Error GoTo fehler
 If frm.anaRS!Hirndurchblutungsstörung = "n" Then ' Hirndurchblutungsstörung
  Call inActi(frm.vTextB(83))  ' Schlaganfall
 Else
  Call Acti(frm.vTextB(83))
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Hirndurchblutungsstörung_Exit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function

#If mitab Then
Function doGilbE(frm As Object, ZielICD$, gilbFeld$, gilbNr%) ' gilb einzeln (ohne kradd)
 If InStrB(frm.vTextB(147), ZielICD) = 0 Then ' Diagnosen
  Call doGilb(frm, ZielICD, gilbFeld, gilbNr)
 End If
End Function ' doGilbE

Function doGilb(frm As Object, ZielICD$, gilbFeld$, gilbNr%)   ' AS Form_Anamnesebogen
  Dim i%, rVorh As New ADODB.Recordset
  On Error GoTo fehler
     If gilbFeld <> vNS Then
      Dim icdlike$ ' replace$(ZielICD, "V", vns) & "[V]?"
      icdlike = REPLACE$(ZielICD, "V", vNS) & "[V]?"
      If icdlike Like "E1*" Or icdlike Like "L89*" Then
       icdlike = Left$(icdlike, InStr(icdlike, ".")) & ".*"
      Else
       icdlike = REPLACE$(ZielICD, "V", vNS) & IIf(LVobMySQL, "[V]?", vNS)
      End If
      If LVobMySQL Then
      ' frm.anaRS!Pat_id
       myFrag rVorh, "SELECT * FROM `fuerdiagexp` WHERE pat_id = " & Pat_ID & " AND icd RLIKE '" & icdlike & "'"
      Else
      ' frm.anaRS!Pat_id
       myFrag rVorh, "SELECT * FROM `fuerdiagexp` WHERE pat_id = " & Pat_ID & " AND (icd LIKE '" & icdlike & "?' OR icd LIKE '" & icdlike & "V?')"
      End If
'      IF rVorh.BOF THEN
'       SET rVorh = Nothing
'       myFrag rVorh, "SELECT * FROM `fuerdiagexp` WHERE pat_id = " & frm.anaRS!Pat_id & " AND icd RLIKE '" & ZielICD & "V'"
'      END IF
''      IF rVorh.NoMatch THEN rVorh.Seek "=", frm.Pat_id, ZielICD + "Z" das wird zu viel
      If rVorh.BOF Then
       On Error Resume Next
       Dim doMarkier%
       doMarkier = False
       If frm.Controls(gilbFeld)(gilbNr).name = "vTextB" Or frm.Controls(gilbFeld)(gilbNr).name = "vCheckb" Then
        doMarkier = True
       Else
        If frm.Controls(gilbFeld)(gilbNr).ControlType <> 106 Then doMarkier = True
       End If
       If doMarkier Then
        frm.Controls(gilbFeld)(gilbNr).BackColor = MarkierFarbe
       End If
       On Error GoTo fehler
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doGilb/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doGilb

Function KRAdd(frm As Form, ByVal flag$, ByVal KrStr$, ZielICD$, DSiAkt As DSiTyp, Optional ICD1$, Optional ICD2$, Optional ICD3$, Optional ICD4$, Optional gilbFeld$, Optional gilbNr%)
  Dim rechts$, i%, obalt%
' Static NKrStr$(), Nflag$(), DSi$() ' für KRAdd
  On Error GoTo fehler
  rechts = vNS
  labPos = 0
  If flag <> vNS And Not IsNull(flag) Then
    If ICD1 = vNS Then ICD1 = ZielICD
    If Not EinRueck(frm, ICD1, KrStr, ICD2, ICD3, ICD4) Then '  AND flag = "Fragebogen") THEN
' Einschub 28.3.09
     obalt = -1
     For i = 1 To UBound(NKrStr)
      If i <= UBound(MDIICD) Then
       If NKrStr(i) = KrStr And MDIICD(i) = ZielICD And DSiAkt <> Zn And DSi(i) <> Zn Then obalt = i: Exit For ' Rezidive zeigen
      End If
     Next i
     If obalt <> -1 Then
      Nflag(obalt) = Nflag(obalt) & "," & flag
      If DSiAkt = gesi Then DSi(obalt) = gesi
     Else
      ReDim Preserve NKrStr(UBound(NKrStr) + 1)
      ReDim Preserve Nflag(UBound(Nflag) + 1)
      ReDim Preserve DSi(UBound(DSi) + 1)
      NKrStr(UBound(NKrStr)) = KrStr
      Nflag(UBound(Nflag)) = flag
      DSi(UBound(DSi)) = DSiAkt
      If MDIn <= UBound(MDIICD) Then
       MDIICD(MDIn) = ZielICD
       MDIDiag(MDIn) = KrStr
       MDIn = MDIn + 1
      End If
     End If
     Call doGilb(frm, ZielICD, gilbFeld, gilbNr)
    
     If rechts = "V" Then
      labPos = 1
     ElseIf rechts <> "Z" And rechts <> "A" Then
      labPos = -1
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in KRAdd/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' kradd(frm,ByVal ob$, KrStr$, Optional icd1$, Optional icd2$)

Function KrStart()
 ReDim NKrStr(0)
 ReDim Nflag(0)
 ReDim DSi(0)
End Function ' KrStart()

Function KrSchluß(frm As AnBog)
 Dim i%
 For i = 1 To UBound(NKrStr)
  If i <= UBound(MDIICD) Then
   If DSi(i) <> gesi Then
   Select Case DSi(i)
    Case Va: frm.vTextB(168) = frm.vTextB(168) & "V.a.":      MDIICD(i) = MDIICD(i) & "V"
    Case Zn: frm.vTextB(168) = frm.vTextB(168) & "Z.n.": If Left$(MDIICD(i), 1) <> "Z" Then MDIICD(i) = MDIICD(i) & "Z"
    Case AuS: frm.vTextB(168) = frm.vTextB(168) & "Ausschl.": MDIICD(i) = MDIICD(i) & "A"
   End Select
   End If 'LenB(DSi(i)) <> 0 THEN
  End If
  frm.vTextB(168) = frm.vTextB(168) & NKrStr(i) & " (" & Nflag(i) & ")" & vbCrLf ' MDi
 Next i '
End Function ' KrSchluß
 
 ' Wenn Laborparameter mit dem Wert "0" nicht wirklich 0 bedeuten ...
Function LabWert0!(frm As Form, name$, Optional sign, Optional grenze!, Optional Diagnose$, Optional ZielICD$, Optional DSiAkt As DSiTyp, Optional Einheit$, Optional ICD1$, Optional ICD2$, Optional SGrenze, Optional WGrenze, Optional ICD3$)  ' sichere Grenze
 Dim source$
 On Error GoTo fehler
 LabWert0 = 0
 'rLa.Seek "=", CStr(Pat_id), name
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", Labwert0: vor find)"
#End If
 If raLau.State = 0 Then ' 2.11.22
  source = raLau.source
  Call DBCnOpen
  Set raLau = Nothing
  raLau.Open source, DBCn, adOpenStatic, adLockReadOnly
 End If
 If Not raLau.BOF Then
  raLau.MoveFirst
  raLau.Find "abkü=" & "'" & name & "'"
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", Labwert0: nach find)"
#End If
  If Not raLau.EOF Then
   If Val(raLau!Wert) = 0 Then
   Else
    LabWert0 = LabWert(frm, name$, sign, grenze!, Diagnose$, ZielICD$, DSiAkt, Einheit$, ICD1$, ICD2$, SGrenze, WGrenze, ICD3$)
   End If
  End If
  raLau.MoveFirst
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabWert0/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LabWert0

Function LabwertA#(frm As Form, laba As LabArt, Optional sign, Optional ByVal grenze!, Optional Diagnose$, Optional ZielICD$, Optional DSiAkt As DSiTyp, Optional Einheit$, Optional ICD1$, Optional ICD2$, Optional SGrenze, Optional WGrenze, Optional ICD3$) ' sichere Grenze, weibliche Grenze
 Dim sq2$, sq1$, Kommentar$, Zeitpunkt As Date
 Dim Labs As labtyp
 On Error GoTo fehler
 alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
 Labs = LabPat(laba, Pat_ID)
 If Labs.Zp <> -1 Then
  LabwertA = MachNumerisch(Labs.WertSg)
  If Alter = 0 Then Alter = AlterBei(Zeitpunkt, GebDat)
  If Not IsMissing(sign) Then
   If Not IsMissing(WGrenze) And obWeib Then grenze = WGrenze
   If Not IsMissing(SGrenze) Then
    If Not ((sign = "<" And LabwertA < SGrenze) Or (sign = ">" And LabwertA > SGrenze)) Then
     DSiAkt = Va
    End If
   End If
   If (sign = "<" And LabwertA < grenze) Or (sign = ">" And LabwertA > grenze) Then
    Call KRAdd(frm, Labs.Abkü & " " & CStr(LabwertA) & " " & Einheit, Diagnose, ZielICD, DSiAkt, ICD1, ICD2, ICD3)
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabWertA/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LabwertA#(frm As Form, LabA As LabArt, Optional sign, Optional ByVal grenze!, Optional Diagnose$, Optional ZielICD$, Optional DSiAkt As DSiTyp, Optional Einheit$, Optional ICD1$, Optional ICD2$, Optional SGrenze, Optional WGrenze, Optional ICD3$) ' sichere Grenze, weibliche Grenze

'Function LabWertD!(frm AS Form, Name$, Optional sign, Optional ByVal Grenze!, Optional Diagnose$, Optional ZielICD$, Optional DSiAkt AS DSiTyp, Optional Einheit$, Optional ICD1$, Optional ICD2$, Optional SGrenze, Optional WGrenze, Optional ICD3$)
' LabWertD = LabWert(frm, Name$, sign, Grenze!, Diagnose$, ZielICD$, DSiAkt, Einheit$, ICD1$, ICD2$, SGrenze, WGrenze, ICD3$)
'End FUNCTION ' LabWertD
' die folgenden Funktionen sind für Fremdlabor und Augenbefunde identisch
Function LabWert!(frm As Form, name$, Optional sign, Optional ByVal grenze!, Optional Diagnose$, Optional ZielICD$, Optional DSiAkt As DSiTyp, Optional Einheit$, Optional ICD1$, Optional ICD2$, Optional SGrenze, Optional WGrenze, Optional ICD3$) ' sichere Grenze, weibliche Grenze
 
 Dim sq2$, sq1$, Wert$, Kommentar$, Zeitpunkt As Date
 On Error GoTo fehler
#If True Then
 Dim rbLau As ADODB.Recordset
 Set rbLau = hollabor(Pat_ID, "^" & name & "$", 0, 0, 0, -1, vNS)
 If Not rbLau.EOF Then
  Wert = rbLau!Wert
  Kommentar = IIf(IsNull(rbLau!Kommentar), "", rbLau!Kommentar)
  Zeitpunkt = rbLau!Zeitpunkt
 End If
#Else
 Dim rbLau As New ADODB.Recordset
' sql = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" AS NB FROM (" & laborAbfr & " WHERE pat_id = " & Pat_id & " AND abkü = '" & Name & "' AND wert <> '') AS labor UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar, Normbereich AS NB " + _
 "FROM `laborxus` LEFT JOIN `laborxwert` ON `laborxus`.RefNr=`laborxwert`.RefNr " & _
 "WHERE pat_id = " & CStr(frm.anaRS!Pat_id) & "  AND abkü = '" & Name & "' AND wert <> '' AND NOT EXISTS (SELECT * FROM `laborneu` WHERE pat_id = " & CStr(frm.anaRS!Pat_id) & " AND abkü = laborxwert.Abkü AND wert = laborxwert.wert AND zeitpunkt > laborxus.Eingang-3 AND zeitpunkt <  laborxus.Eingang+6) "
' sql = "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_id & " AND abkü = '" & Name & "' AND wert <> '' UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_id & " AND abkü = '" & Name & "' AND wert <> '') i GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb ORDER BY zeitpunkt DESC"
 sq2 = "SELECT * FROM `labor2a` WHERE pat_id = " & Pat_ID & " AND abkü = '" & name & "' AND wert <> '' ORDER BY zeitpunkt DESC" ' GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb
 sq1 = "SELECT * FROM `labor1a`  WHERE pat_id = " & Pat_ID & " AND abkü = '" & name & "' AND wert <> '' ORDER BY zeitpunkt DESC" ' GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb
 
' SET rblau = Dtb.OpenRecordset(sql)
 Set rbLau = Nothing
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", Labwert: vor sql)"
#End If
' rbLau.Open sq2, DBCn, adOpenDynamic, adLockReadOnly
 myFrag rbLau, sq2
 If Not rbLau.EOF Then
  Wert = rbLau!Wert
  Kommentar = IIf(IsNull(rbLau!Kommentar), "", rbLau!Kommentar)
  Zeitpunkt = rbLau!Zeitpunkt
 End If
 Set rbLau = Nothing
#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", Labwert: zwischen sql)"
#End If
' rbLau.Open sq1, DBCn, adOpenDynamic, adLockReadOnly
 myFrag rbLau, sql
 If Not rbLau.EOF Then
  If rbLau!Zeitpunkt > Zeitpunkt Then
   Wert = rbLau!Wert
   Kommentar = IIf(IsNull(rbLau!Kommentar), "", rbLau!Kommentar)
   Zeitpunkt = rbLau!Zeitpunkt
  End If
 End If
 Set rbLau = Nothing

#If debu <> 0 Then
       Tvor = Takt: Takt = Timer
       dnr = dnr + 1
       Print #313, Format(Takt - Tvor, "0.00") & "      " & Format(Takt - T0, "0.00") & " (" & dnr & ", Labwert: nach sql)"
#End If
#End If ' true else
 LabWert = 0
 'rlau.Seek "=", CStr(Pat_id), name
' rblau.Find "abkü = " & "" & Name & "" & " AND wert <> """"", 0, adSearchForward, 1
 If Zeitpunkt <> 0 Then
  LabWert = Val(REPLACE$(REPLACE$(IIf(IsNull(Wert), IIf(IsNull(Kommentar), vNS, Kommentar), Wert), ",", "."), "%", vNS))
'  IF Alter = 0 THEN Alter = (Zeitpunkt - GebDat) * 2.73792574745373E-03 ' 1/365,24
  If Alter = 0 Then Alter = AlterBei(Zeitpunkt, GebDat)
  If Not IsMissing(sign) Then
   If Not IsMissing(WGrenze) And obWeib Then grenze = WGrenze
   If Not IsMissing(SGrenze) Then
    If Not ((sign = "<" And LabWert < SGrenze) Or (sign = ">" And LabWert > SGrenze)) Then
      DSiAkt = Va
    End If
   End If
   If (sign = "<" And LabWert < grenze) Or (sign = ">" And LabWert > grenze) Then
    Call KRAdd(frm, UCase$(name) & " " & CStr(LabWert) & " " & Einheit, Diagnose, ZielICD, DSiAkt, ICD1, ICD2, ICD3)
   End If
  End If
 End If
' rlau.MoveFirst
' Do
'  IF rlau!Pat_id = Pat_id THEN
'   Debug.Print rlau!Abkü, rlau!Wert
'  END IF
'  rlau.MoveNext
' Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LabWert/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LabWert
#End If

Function EinRueck%(frm As Form, ICD1$, Diagnose$, Optional ICD2$, Optional ICD3$, Optional ICD4$)
  On Error GoTo fehler
' 15.6.: anam!diagnosen statt anam!diagnosen
  If ICD1 = vNS Then Exit Function
'  ON Error Resume Next
  If InStrB(frm.vTextB(147), ICD1) = 0 Then
   If ICD2 = vNS Then Exit Function
   If ICD2 = ICD1 Then Exit Function
   If InStrB(frm.vTextB(147), ICD2) = 0 Then
    If ICD3 = vNS Then Exit Function
    If ICD3 = ICD2 Or ICD3 = ICD1 Then Exit Function
    If InStrB(frm.vTextB(147), ICD3) = 0 Then
     If ICD4 = vNS Then Exit Function
     If ICD4 = ICD3 Or ICD4 = ICD2 Or ICD4 = ICD1 Then Exit Function
     If InStrB(frm.vTextB(147), ICD4) = 0 Then Exit Function
    End If
   End If
  End If
  Diagnose = "  " + Diagnose
  EinRueck = -1
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in EinRueck/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Einrück

#If wennsvorkommt Then
Function do_Fremdlabor_Form_Current(frm As Form)
  Dim dFrs As ADODB.Recordset
  On Error GoTo fehler
'  IF rsAnam Is Nothing THEN
'   Call do_Fremdlabor_Form_Load(frm)
'  END IF
'  rsNa.Seek "=", frm.anaRS!Pat_id
'  Set rsNa = Nothing
  ' frm.anaRS!Pat_id
  myFrag dFrs, "SELECT * FROM `namen` WHERE pat_id = " & Pat_ID
  If Not dFrs.EOF Then
   frm.anaRS!PatName = GesNamFn(dFrs)  ' rNa!Nachname & " " & rNa!Vorname
  End If
  Set dFrs = Nothing
  If Not IsNull(frm.anaRS!DokPfad) Then
'   dFrs.Seek "=", CLng(frm.anaRS!Pat_id), CStr(frm.anaRS!DokPfad)
'   IF Not dFrs.NoMatch THEN
' frm.anaRS!Pat_id
'   dFrs.Open "SELECT * FROM `dokumente` WHERE pat_id = " & Pat_id & " AND dokpfad = '" & frm.anaRS!DokPfad & "'", DBCn, adOpenDynamic, adLockReadOnly
   myFrag dFrs, "SELECT * FROM briefe WHERE pat_id = " & Pat_ID & " AND pfad = '" & frm.anaRS!DokPfad & "'"
   If Not dFrs.BOF Then
    frm.anaRS!Eintragsdatum = dFrs!Zeitpunkt
   End If ' dFrs.NoMatch THEN
   Set dFrs = Nothing
   Call makeDatPfad(frm)
   frm.Controls![gescannte Datei].Caption = "&Zeig: " + frm.anaRS!DokName
  Else ' ISNULL(frm.anaRS!DokPfad) THEN
   DatPfad = vNS
   frm.Controls![gescannte Datei].Caption = vNS
  End If ' ISNULL(frm.anaRS!DokPfad) THEN
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Fremdlabor_Form_Current/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_Fremdlabor_Form_Current(frm As Form)
#End If

Function makeDatPfad(frm As Form)
 On Error GoTo fehler
 If (InStrB("$\", Left$(frm!DokPfad, 1)) <> 0 And Left$(frm!DokPfad, 1) <> vNS) Or Mid$(frm!DokPfad, 2, 1) = ":" Then
   RDatPfad = REPLACE$(REPLACE$(LCase$(IIf(IsNull(frm!DokPfad), vNS, frm!DokPfad)), "$\turbomed\dokumente", getDokPfad), "^", vNS)
   DatPfad = vNS + RDatPfad + vNS
   Dim cat As New ADOX.Catalog
   Dim cmd As New ADODB.Command
   cat.ActiveConnection = DBCn
   Set cmd = cat.Views("LaborDokumente ep").Command
   If InStrB(cmd.CommandText, "FROM `" + QMdbAkt + "`.`briefe`") <> 0 Then
'   IF instrb(Dtb.QueryDefs![LaborDokumente ep].sql, "FROM `" + QMdbAkt + "`.`briefe`") <> 0 THEN
    Dim Appl$, Bef$
    Appl = getReg(0, ".doc", vNS)
    Bef = getReg(0, Appl + "\shell\open\command", vNS)
'    DatPfadPI = Bef & " " & DatPfad '"c:\programme\microsoft office\office10\winword.exe " + DatPfad
    DatPfadPI = "winword " + DatPfad
    Bef = REPLACE$(REPLACE$(getReg(0, "wordpad.document.1\shell\open\command", vNS), "%1", vNS), "%programfiles%", Environ("ProgramFiles"))
    DatPfad = Bef & " " & DatPfad '"c:\programme\windows nt\zubehör\wordpad.exe " + DatPfad
    frm.Anzeigen.Caption = "Word&Pad"
    frm.Photoimpact.Caption = "&Winword"
   Else
    DatPfadPI = Environ("ProgramFiles") & "\Ulead Systems\Ulead PhotoImpact 6\iedit.exe " + DatPfad
    DatPfad = getIViewPfad & " " & DatPfad ' Environ("ProgramFiles") & "\irfanview\i_view32.exe
   End If
  frm!Anzeigen.Enabled = -1
  frm!Photoimpact.Enabled = -1
  frm!DokName.Enabled = 0
 Else
  DatPfad = vNS
  DatPfadPI = vNS
  frm!Anzeigen.Enabled = 0
  frm!Photoimpact.Enabled = 0
  frm!DokName.Enabled = -1
 End If
 
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in makeDatPfad/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' makeDatPfad

' 4.9.06: wird nirgends aufgerufen
Function do_LaborDokumente_Form_Current(frm As Form)
 Static fS, F
 On Error GoTo fehler
 Call makeDatPfad(frm)
'frm.Controls!Nachname
' rDokab.Seek "=", frm!DokPfad
' IF Not rDokab.NoMatch THEN
'   frm!abgehakt = rDokab!abgehakt
'END IF
' frm!DokGroe = CreateObject("Scripting.FileSystemObject").getfile(replace$(replace$(IIF(ISNULL(frm!DokPfad), vns, frm!DokPfad), "$\TurboMed\Dokumente\", getDokPfad), "^", vns)).size
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_LaborDokumente_Form_Current/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_LaborDokumente_Form_Current(frm As Form)

Function do_LaborDokumente_form_load(frm As Form)
 On Error GoTo fehler
' SET raDokab = TabÖff("br_abgehakt", "DokPfad")
 Dim raDokab As ADODB.Recordset ' Abgehakte Dokumente
' Set raDokab = Nothing
' Call raDokab.Open("SELECT -ob AS j_ob, d.* FROM `br_abgehakt` d", DBCn, adOpenDynamic, adLockReadOnly)
 myFrag raDokab, "SELECT -ob j_ob, d.* FROM `br_abgehakt` d"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_LaborDokumente_form_load/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_LaborDokumente_form_load(frm As Form)

Public Sub Taste(KeyCode%, Shift%, frm As Form)
On Error GoTo fehler
    If KeyCode = 40 Then
        DoCmd.GoToRecord , , acNext
    ElseIf KeyCode = 38 Then
        DoCmd.GoToRecord , , acPrevious
    ElseIf KeyCode = 27 Then
        DoCmd.Close acForm, frm.name, acSaveYes
    End If
Ende:
    Exit Sub
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Taste/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub

Function do_Fremdlabor_Form_Load(frm As Form)
' ON Error GoTo fehler
' SET rDok = TabÖff("Dokumente", "PIDokPfad")
' SET rsNa = TabÖff("Namen", "pat_ID")
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIF(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Fremdlabor_Form_Load/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
End Function
Function do_Fremdlabor_Form_Unload(frm As Form, Cancel%)
' ON Error GoTo fehler
' rDok.Close
' rsNa.Close
' rDokab.Close
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIF(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_Fremdlabor_Form_Unload/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
End Function

Function do_gescannte_Datei_Click(frm As Form)
 On Error GoTo fehler
 Call Shell(DatPfad, vbMaximizedFocus)
 frm!Datum.SetFocus
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_gescannte_Datei_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_gescannte_Datei_Click

Function do_anzeigen_click(frm As Form)
 On Error GoTo fehler
 If InStrB(DatPfad, ".exe") <> 0 Then
  Call Shell(DatPfad, vbMaximizedFocus)
 Else
  If Left$(DatPfad, InStr(DatPfad, " ")) = "winword" Then
   GetWord
   Wapp.Visible = True
   Wapp.WindowState = wdWindowStateMaximize
   Wapp.documents.Open DatPfad
   Wapp.Activate
  Else
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_anzeigen_click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_anzeigen_click

Function do_PhotoImpact_Click(frm As Form)
 On Error GoTo fehler
 If InStrB(DatPfadPI, ".exe") <> 0 Then
  Call Shell(DatPfadPI, vbMaximizedFocus)
 Else
  If Left$(DatPfadPI, InStr(DatPfadPI, " ") - 1) = "winword" Then
   GetWord
   Wapp.Visible = True
   Wapp.WindowState = wdWindowStateMaximize
   Wapp.documents.Open Mid$(DatPfadPI, InStrB(DatPfadPI, " "))
   Wapp.Activate
  Else
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_PhotoImpact_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_PhotoImpact_Click

Function do_FDatenquelle_Change(frm As Form)
Dim Pat_ID&
 On Error GoTo fehler
Pat_ID = frm.Recordset!Pat_ID
On Error Resume Next ' "nicht zulässsig" kam einmal
frm.RecordSource = frm!FDatenquelle
On Error GoTo fehler
frm.Recordset.FindFirst "Pat_id = " + CStr(Pat_ID)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_FDatenquelle_Change/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_FDatenquelle_Change(frm AS Form)

'Function ListFreigaben()
'Dim Nr&, erg$, Inhalt$
'Nr = 0
'FZ = 0
'Do
'' erg = ReadRegistryGetValues(&H80000002, "SYSTEM\ControlSet001\Services\lanmanserver\Shares", nr, Inhalt)
' erg = ReadRegistryGetValues(&H80000002, "SYSTEM\CurrentControlSet\Services\lanmanserver\Shares", Nr, Inhalt)
' IF Inhalt <> vns THEN
'  IF mid$(Inhalt, 2, 1) = ":" THEN
'   ReDim Preserve FNam(FZ)
'   ReDim Preserve FInh(FZ)
'   FNam(FZ) = erg
'   FInh(FZ) = Inhalt
'   FZ = FZ + 1
'   Debug.Print erg
'   Debug.Print Inhalt
'  END IF
' END IF
' Nr = Nr + 1
'Loop Until trim$(erg) = vns
'End Function
'function ReadRegistryGetValues$(ByVal Group&, ByVal Section$, Idx&, Inhalt$)
'Dim lResult&, lKeyValue&, lValueLength&, sValue$, td#, res&, vTyp&
'Dim ValueN$, cbValueN&
'ValueN = Space$(512)
'cbValueN = Len(ValueN)
'sValue = Space$(512)
'lValueLength = Len(sValue)
'Inhalt = vns
'sValue = vns
'On Error Resume Next
'lResult = RegOpenKeyEx(Group, Section, 0, &HF003F, lKeyValue)
'If lResult = 0 THEN
' lResult = RegEnumValue(lKeyValue, Idx, ValueN, cbValueN, 0&, vTyp, sValue, lValueLength)
' IF (lResult = 0) AND (eRR.Number = 0) THEN
'   sValue = LEFT(sValue, InStr(sValue, vbnullchar) - 1)
'   ValueN = LEFT(ValueN, InStr(ValueN, vbnullchar) - 1)
'   ReadRegistryGetValues = ValueN
'   IF (instrb(sValue, "Path") <> 0) THEN
'    Inhalt = mid$(sValue, InStr(sValue, "Path") + 5)
'    Inhalt = LEFT(Inhalt, InStr(Inhalt, "Permiss") - 2)
'   END IF
'   lResult = RegCloseKey(lKeyValue)
' END IF
'
'END IF
'End FUNCTION ' ReadRegistryGetSubkey$(ByVal Group&, ByVal Section$, Idx&)
Function ReadRegistryGetSubkey$(ByVal Group&, ByVal Section$, idx&)
 Dim lResult&, lKeyValue&, lDataTypeValue&, lValueLength&, sValue$, td#
 On Error GoTo fehler
 sValue = Space$(2048)
 lValueLength = Len(sValue)
 On Error Resume Next
 lResult = RegOpenKeyEx(Group, Section, 0, &HF003F, lKeyValue)
 If lResult = 0 Then
  lResult = RegEnumKey(lKeyValue, idx, sValue, lValueLength)
  If (lResult = 0) And (Err.Number = 0) Then
   sValue = Left$(sValue, InStr(sValue, vbNullChar) - 1)
  End If
  ReadRegistryGetSubkey = sValue
  lResult = RegCloseKey(lKeyValue)
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ReadRegistryGetSubkey/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ReadRegistryGetSubkey$(ByVal Group&, ByVal Section$, Idx&)


Sub GetExcel()
    Dim XL1 As Excel.Application 'Object    ' Variable für Verweis auf
                            ' Microsoft Excel.
    Dim ExcelLiefNicht As Boolean    ' Attribut für Freigabe am Ende.

' Überprüfen, ob eine Kopie von Microsoft Excel bereits
' ausgeführt wird.
On Error Resume Next    ' Fehlerbehandlung zurückstellen.
' GetObject-Funktionsaufruf ohne erstes Argument gibt einen Verweis auf
' eine Instanz der Anwendung zurück. Wenn die Anwendung nicht
' ausgeführt wird, tritt ein Fehler auf.
    Set XL1 = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then ExcelLiefNicht = True
    Err.Clear    ' Err-Objekt im Fehlerfall löschen.

' Prüfen auf Microsoft Excel. Wenn Microsoft Excel ausgeführt wird, wird
' dies in die Tabelle ausgeführter Objekte eingetragen.
    DetectExcel

' Objektvariable so festlegen, daß sie auf die gewünschte Datei verweist.
    Set XL1 = GetObject("c:\vb4\TEST1.XLS")

' Microsoft Excel mit zugehöriger Application-Eigenschaft einblenden.
' Fenster mit der Datei unter Verwendung der Windows-Auflistung des
' XL1-Objektverweises anzeigen.
    XL1.Application.Visible = True
    XL1.Parent.Windows(1).Visible = True
' Dateiverarbeitung.
    ' ...
' Wenn diese Kopie von Microsoft Excel beim Starten des Beispiels
' nicht ausgeführt wurde, wird Excel mit der Quit-Methode des
' Application-Objekts beendet. Wenn Sie versuchen, Microsoft Excel zu
' beenden, blinkt die Titelleiste, und Sie werden
' in einer Meldung gefragt, ob Sie geladene Dateien speichern möchten.
    If ExcelLiefNicht = True Then
        XL1.Application.Quit
    End If

    Set XL1 = Nothing    ' Verweis auf Anwendung und
                            ' Tabelle freigeben.
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetExcel/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' GetExcel

Sub DetectExcel()
' Diese Prozedur erkennt, ob Excel ausgeführt wird, und registriert dies.
    Const WM_USER = 1024
    Dim hwnd&
' Wenn Excel ausgeführt wird, wird durch Ausführen dieses
' API-Aufrufs die zugehörige Zugriffsnummer zurückgegeben.
    On Error GoTo fehler
    hwnd = FindWindow("XLMAIN", 0)
    If hwnd = 0 Then    ' 0 bedeutet, daß Excel nicht ausgeführt wird.
        Exit Sub
    Else
    ' Excel wird ausgeführt. Verwenden der API-Funktion
    ' SendMessage, um Excel in der Tabelle ausgeführter
    ' Objekte einzutragen.
        SendMessage hwnd, WM_USER + 18, 0, 0
    End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DetectExcel/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' DetectExcel

Public Sub Detectword(className$)
' Procedure detects a running Word AND r
' egisters it.
 Const WM_USER = 1024
 Dim hwnd&
' IF Word is running this API call retur
' ns its handle.
 On Error GoTo fehler
 hwnd = FindWindow(className, vNS)

 If hwnd = 0 Then ' 0 means Word Not running.
 Exit Sub
 Else
' Word is running so use the SendMessage
' API FUNCTION to enter it in the Running
' Object Table.
 SendMessage hwnd, WM_USER + 18, 0, 0
 End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Detectword/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Detectword

Private Sub cmdPreview_Click(myfilepath$)
'Preview a word document
On Error GoTo fehler
Screen.MousePointer = 11 'vbHourglass
'Startup Word IF not started, OR switch to existing one
GetWord
'Open the document, maximised, AND switch to word
Wapp.Visible = True
Wapp.Application.WindowState = wdWindowStateMaximize
Wapp.documents.Open (myfilepath)
Screen.MousePointer = 0 ' vbDefault
Wapp.Application.Activate
Set Wapp = Nothing
'We can also use the 'WordWasNotRunning' flag to decide IF we should quit word OR not
'Users might get annoyed IF docs they are currently editing aren't saved!!
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in cmdPreview_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' cmdPreview_Click

#If mitab Then
Function do_DMPAusgeb1(frm As Form)
' frm.anaRS!Pat_id
 Call do_DMPAusgebStandAlone(Pat_ID)
End Function ' do_DMPAusgeb1
#End If

Sub FaxSend(docName$, RecName$, RecNum$)
Dim FaxServer As New FAXCOMLib.FaxServer
Dim FaxDoc As New FAXCOMLib.FaxDoc
Dim FaxTiff As New FAXCOMLib.FaxTiff
Dim strFaxJob As FAXCOMLib.FaxJobs
Dim strFaxStatus As FAXCOMLib.FaxJob
Dim DateiName$
Dim strFaxTiff As FAXCOMLib.FaxTiff
Dim strJobID&
On Error GoTo fehler
 DateiName = FaxSendDatei
 Err.Clear
 FaxServer.connect ("anmeld2")

Set FaxDoc = FaxServer.CreateDocument(DateiName)
   
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

Function tha12()
 Dim h$()
 Call Lese.ProgStart
 Call getHausarzt(1576, h)
 Call Lese.ProgEnde
End Function ' tha12

' in einDMP, dodoFollowUp, doUngeschriebeneBriefe, PatAuswahl.do_Pat_ID_Change, tha12
Function getHausarzt(pid&, infos$()) ' Bildet aus den Infos in `namen` und `hareal` die alte gleichnamige Funktion nach, die jetzt noch unter getHausarztAlt zur Verfügung steht
' 5.12.12: 15. Feld briefmail

' SELECT pat_id pid, if(h.anrede,"Herr","Frau"), Adressat, h.Straße, PLZOrt, Fax, Überschrift, if(dmp2,'X','') dmp2,
' if(dmp1,'X','') dmp1, Niederlassungsgebiet, h.Vorname, Funktion, if(InnereAllg,'X','') InnereAllg, h.kvnr, Tel, h.Nachname
' FROM `namen`
' LEFT JOIN hareal h ON h.kvnr IN (getHA0,getHA1,getHA2) AND h.kvnr<>0
' Where Pat_id = pid
' ORDER BY pat_id,if(h.kvnr=getha0,0,if(h.kvnr=getha1,1,2))
' ;

 Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset, i&, j&
 Dim HA&(2), fnHA$(2)
 On Error GoTo fehler
 myFrag rs, "SELECT getHA0, fnHA0, getHA1, fnHA1, getHA2, fnHA2 FROM `namen` WHERE pat_id = " & pid
 If Not rs.EOF Then
  If Not IsNull(rs!getHA0) Then HA(0) = rs!getHA0: fnHA(0) = rs!fnHA0
  If Not IsNull(rs!getHA1) Then HA(1) = rs!getHA1: fnHA(1) = rs!fnHA1
  If Not IsNull(rs!getHA2) Then HA(2) = rs!getHA2: fnHA(2) = rs!fnHA2
  If HA(0) = 0 Then
   If HA(1) <> 0 Then
    HA(0) = HA(1)
    HA(1) = 0
    fnHA(0) = fnHA(1)
    fnHA(1) = ""
   ElseIf HA(2) <> 0 Then
    HA(0) = HA(2)
    fnHA(0) = fnHA(2)
    HA(2) = 0
    fnHA(2) = ""
   End If
  End If
  If HA(0) <> 0 Then
   If HA(1) <> 0 Then
    If HA(2) <> 0 Then
     ReDim infos(15, 2)
    Else
     ReDim infos(15, 1)
    End If
   Else
    ReDim infos(15, 0)
   End If
   For i = 0 To UBound(infos, 2)
    Set rs1 = Nothing
    If HA(i) <> 0 Then
     myFrag rs1, "SELECT * FROM `hareal` WHERE kvnr = " & HA(i)
     If Not rs1.EOF Then
      For j = 0 To 15
       If j <> 10 Then
       Select Case j
        Case 0
         infos(j, i) = IIf(rs1.Fields(j) = 0, "Frau", "Herr")
        Case 6, 7, 11
         infos(j, i) = IIf(rs1.Fields(j) = 0, vNS, "X")
        Case Else
         On Error Resume Next
        
         infos(j, i) = rs1.Fields(j)
         On Error GoTo fehler
       End Select ' j
       End If ' j <> 10
      Next j
      infos(10, i) = fnHA(i)
     End If ' Not rs1.EOF THEN
    End If ' NOT ISNULL
   Next i ' i = 0 To UBound(Infos, 2)
  Else ' NOT ISNULL(HA0) AND HA0 <> 0 THEN
   ReDim infos(15, 0)
  End If ' NOT ISNULL(HA0) AND HA0 <> 0 THEN
 End If ' Not rs.EOF THEN
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getHausarzt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getHausarzt(Pid&, Infos$())

'Function getHausarztAlt(Pid&, Infos$(), Optional obHAPrio%)
'' Dim rNa As New ADODB.Recordset
' Dim rHa As New ADODB.Recordset, rAF&
' Dim gefunden%, runde%, irunde%, HACngef%
' Dim rListena As New ADODB.Recordset
' Dim rHae As New ADODB.Recordset
' Dim faxfehlt%
'' Dim rFa As New ADODB.Recordset
'' Dim i%, j%
' Dim InfRoh$()
' Dim üwerg$() ' HaErg$()
' ON Error GoTo fehler
' ReDim Infos(14, 2)
' '  0: Frau/Herrn
' '  1: Titel+Vorn+Nachn,
' '  2: Straße,
' '  3: PLZ+Ort,
' '  4: Faxnr,
' '  5: S.g./Liebe,
' '  6: DMPTyp2,
' '  7: DMPTyp1,
' '  8: Niederlassungsgebiet 3. Feld für einmal austauschen,
' '  9: Vorname,
' ' 10: Funktion ("Üw 207, HA"),
' ' 11: Fachrichtung
' ' 12: KV-Nummer
' ' 13: Tel'nr.
' ' 14: Nachname,
'
' Call acon(HaT)
'
'' in Faelle steht der als Überweiser eingetragene
' IF Üw12(Pid, üwerg) = vNS THEN Exit Function
'
'    ReDim InfRoh(14, UBound(üwerg, 2))
'    ReDim Infos(14, UBound(üwerg, 2))
'    For runde = 0 To UBound(üwerg, 2)
'     IF üwerg(0, runde) <> vNS THEN
'
'      gefunden = 0
'      IF üwerg(1, runde) <> vNS AND üwerg(2, runde) <> vNS THEN
'       SET rListena = Nothing
'       myFrag rListena, "SELECT * FROM `liuez` WHERE LEFT(kvnr,7) = '" & Left$(üwerg(0, runde), 7) & "' AND name = '" & üwerg(1, runde) & "' AND vorname = '" & üwerg(2, runde) & "' ORDER BY fax DESC"
'       IF Not rListena.BOF THEN gefunden = True
'      END IF
'      IF Not gefunden AND üwerg(1, runde) <> vNS THEN
'       SET rListena = Nothing
'       myFrag rListena, "SELECT * FROM `liuez` WHERE LEFT(kvnr,7) = '" & Left$(üwerg(0, runde), 7) & "' AND name = '" & üwerg(1, runde) & "' ORDER BY fax DESC"
'       IF Not rListena.BOF THEN gefunden = True
'      END IF
'      IF Not gefunden THEN
'       SET rListena = Nothing
'       myFrag rListena, "SELECT * FROM `liuez` WHERE LEFT(kvnr,7) = '" & Left$(üwerg(0, runde), 7) & "' ORDER BY fax DESC"
'       IF Not rListena.BOF THEN gefunden = True
'      END IF
'      IF Not gefunden THEN
'       SET rListena = Nothing
'       myFrag rListena, "SELECT * FROM `liuez` WHERE name = '" & üwerg(1, runde) & "' AND vorname = '" & üwerg(2, runde) & "' ORDER BY fax DESC"
'       IF Not rListena.BOF THEN gefunden = True
'      END IF
'      IF Not gefunden THEN
'       SET rListena = Nothing
'       Dim sql$
'       sql = "SELECT titel, Niederlassungsart arzttyp, GROUP_CONCAT(DISTINCT a2.nachname) gemmit, fr.fachrichtung fachgruppe, mail email, LEFT(bsnr,7) kvnr, fax, REPLACE(GROUP_CONCAT(DISTINCT tel),'-','') tel, REPLACE(GROUP_CONCAT(DISTINCT fax),'-','') fax, a.nachname name, a.vorname, bs.straße strasse, bs.plz, ort, IF(a.obweibl=1,'w','m') geschlecht, IF(a.obweibl=1,'Frau','Herrn') anrede, tel telefon, GROUP_CONCAT(DISTINCT fachrichtung) zulg, " & _
'           "MAX(IF(ISNULL((SELECT genehmigung FROM " & hadbname & ".arzt_has_genehmigung ahg1 LEFT JOIN " & hadbname & ".genehmigung g1 ON g1.idgenehmigung = ahg1.genehmigung_id WHERE ahg1.arzt_id = a2.idarzt AND genehmigung = 'DMP-DM1_Koordinierender Arzt_Hausarzt')),'0','1')) dmpt1, " & _
'           "MAX(IF(ISNULL((SELECT genehmigung FROM " & hadbname & ".arzt_has_genehmigung ahg2 LEFT JOIN " & hadbname & ".genehmigung g2 ON g2.idgenehmigung = ahg2.genehmigung_id WHERE ahg2.arzt_id = a2.idarzt AND genehmigung = 'DMP-DM2_Koordinierender Arzt')),'0','1')) dmpt2 " & _
'           "FROM " & hadbname & ".arzt a " & _
'           "LEFT JOIN " & hadbname & ".arzt_has_fachrichtung ahf ON a.idarzt = ahf.arzt_id LEFT JOIN " & hadbname & ".fachrichtung fr ON ahf.fachrichtung_id = fr.idfachrichtung " & _
'           "LEFT JOIN " & hadbname & ".arzt_has_bs ahb ON a.idarzt = ahb.arzt_id LEFT JOIN " & hadbname & ".bs ON bs.idbs = ahb.bs_id " & _
'           "LEFT JOIN " & hadbname & ".tel ON tel.bs_id = bs.idbs " & _
'           "LEFT JOIN " & hadbname & ".fax ON fax.bs_id = bs.idbs " & _
'           "LEFT JOIN " & hadbname & ".mail ON mail.bs_id = bs.idbs " & _
'           "LEFT JOIN " & hadbname & ".arzt_has_bs ahb2 ON bs.idbs = ahb2.bs_id LEFT JOIN " & hadbname & ".arzt a2 ON a2.idarzt = ahb2.arzt_id " & _
'           "LEFT JOIN " & hadbname & ".nlart ON a.nlart_id = nlart.idnlart " & _
'           "LEFT JOIN " & hadbname & ".titel t ON t.idtitel = a.titel_id " & _
'           "LEFT JOIN " & hadbname & ".ort o ON o.idort = bs.ort_id " & _
'           "WHERE LEFT(bsnr,7) = '" & Left$(üwerg(0, runde), 7) & "'"
'       myFrag rListena, sql
'       IF Not rListena.BOF THEN gefunden = True
'      END IF
'      IF gefunden THEN
'       InfRoh(1, runde) = IIF(ISNULL(rListena!titel) OR LenB(rListena!titel) = 0, "Dr.med.", rListena!titel) & " " & rListena!vorname & " " & rListena!Name
'       IF NOT ISNULL(rListena!strasse) THEN
'        InfRoh(2, runde) = rListena!strasse
'        InfRoh(3, runde) = rListena!plz & " " & rListena!ort
'        IF ISNULL(rListena!Fax) THEN
'         Dim rfax As New ADODB.Recordset
'         SET rfax = Nothing
'         myFrag rfax, "SELECT faxzahl FROM " & hadbname & ".bs bsr LEFT JOIN " & hadbname & ".arzt_has_bs ahbr ON bsr.idbs = ahbr.bs_id LEFT JOIN " & hadbname & ".arzt a ON a.idarzt = ahbr.arzt_id LEFT JOIN " & hadbname & ".arzt_has_bs ahb ON a.idarzt = ahb.arzt_id LEFT JOIN " & hadbname & ".bs ON bs.idbs = ahb.bs_id LEFT JOIN " & hadbname & ".fax ON fax.bs_id = bsr.idbs WHERE LEFT(bs.bsnr,7) = '" & Left$(üwerg(0, runde), 7) & "' AND NOT ISNULL(faxzahl)"
'         IF Not rfax.BOF THEN
'          InfRoh(4, runde) = rfax!faxzahl
'         END IF
'        Else
'         InfRoh(4, runde) = rListena!Fax
'        END IF
'        InfRoh(12, runde) = rListena!KVNr
'       END IF
'       IF ISNULL(rListena!fachgruppe) THEN InfRoh(8, runde) = vNS ELSE InfRoh(8, runde) = rListena!fachgruppe
'       IF ISNULL(rListena!vorname) THEN InfRoh(9, runde) = vNS ELSE InfRoh(9, runde) = rListena!vorname
'       IF ISNULL(rListena!Name) THEN InfRoh(14, runde) = vNS ELSE InfRoh(14, runde) = rListena!Name
'       InfRoh(10, runde) = üwerg(3, runde)
'       IF ISNULL(rListena!telefon) THEN
'         SET rfax = Nothing
'         myFrag rfax, "SELECT tel FROM " & hadbname & ".bs bsr LEFT JOIN " & hadbname & ".arzt_has_bs ahbr ON bsr.idbs = ahbr.bs_id LEFT JOIN " & hadbname & ".arzt a ON a.idarzt = ahbr.arzt_id LEFT JOIN " & hadbname & ".arzt_has_bs ahb ON a.idarzt = ahb.arzt_id LEFT JOIN " & hadbname & ".bs ON bs.idbs = ahb.bs_id LEFT JOIN " & hadbname & ".tel ON tel.bs_id = bsr.idbs WHERE LEFT(bs.bsnr,7) = '" & Left$(üwerg(0, runde), 7) & "' AND NOT ISNULL(tel)"
'         IF Not rfax.BOF THEN
'          InfRoh(13, runde) = rfax!tel
'         END IF
'       Else
'        InfRoh(13, runde) = rListena!telefon
'       END IF
'      END IF
'      IF VarType(üwerg(0, runde)) = vbString THEN
'       üwerg(0, runde) = CDbl(replace$(üwerg(0, runde), Space$(1), vNS))
'      END IF
'      IF üwerg(0, runde) < 9999999 THEN üwerg(0, runde) = üwerg(0, runde) * 100
'      Dim rDMP As New ADODB.Recordset
'      SET rDMP = Nothing
''      rDMP.Open "SELECT genehmigung FROM (" & hadbname & ".arzt a LEFT JOIN " & hadbname & ".arzt_has_bs ahb ON ahb.arzt_id = a.idarzt)  LEFT JOIN " & hadbname & ".bs ON ahb.bs_id = bs.idbs LEFT JOIN " & hadbname & ".arzt_has_bs ahb2 ON bs.idbs = ahb2.bs_id LEFT JOIN " & hadbname & ".arzt a2 ON ahb2.arzt_id = a2.idarzt LEFT JOIN " & hadbname & ".arzt_has_genehmigung ahg ON ahg.arzt_id = a2.idarzt LEFT JOIN " & hadbname & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE bs.bsnr = '" & üwerg(0, runde) & "' AND genehmigung IN ('DMP-DM2_Koordinierender Arzt','DMP-DM2_FA_Diab.bes.qualif.Arzt(SPP)+Koordin. Arzt')", HAECn, adOpenStatic, adLockReadOnly
'      rDMP.Open "SELECT genehmigung FROM " & hadbname & ".bs LEFT JOIN " & hadbname & ".arzt_has_bs ahb2 ON bs.idbs = ahb2.bs_id LEFT JOIN " & hadbname & ".arzt a2 ON ahb2.arzt_id = a2.idarzt LEFT JOIN " & hadbname & ".arzt_has_genehmigung ahg ON ahg.arzt_id = a2.idarzt LEFT JOIN " & hadbname & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE LEFT(bs.bsnr,7) = '" & Left$(üwerg(0, runde), 7) & "' AND genehmigung IN ('DMP-DM2_Koordinierender Arzt','DMP-DM2_FA_Diab.bes.qualif.Arzt(SPP)+Koordin. Arzt')", HAECn, adOpenStatic, adLockReadOnly
'      InfRoh(6, runde) = IIF(rDMP.BOF, vNS, "X")
'      SET rDMP = Nothing
''      rDMP.Open "SELECT genehmigung FROM (" & hadbname & ".arzt a LEFT JOIN " & hadbname & ".arzt_has_bs ahb ON ahb.arzt_id = a.idarzt)  LEFT JOIN " & hadbname & ".bs ON ahb.bs_id = bs.idbs LEFT JOIN " & hadbname & ".arzt_has_bs ahb2 ON bs.idbs = ahb2.bs_id LEFT JOIN " & hadbname & ".arzt a2 ON ahb2.arzt_id = a2.idarzt LEFT JOIN " & hadbname & ".arzt_has_genehmigung ahg ON ahg.arzt_id = a2.idarzt LEFT JOIN " & hadbname & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE bs.bsnr = '" & üwerg(0, runde) & "' AND genehmigung IN ('DMP-DM1_Koordinierender Arzt(SPP)_Erwachsene','DMP-DM1_Koordinierender Arzt(SPP)_Kinder/Jugendliche','DMP-DM1_Koordinierender Arzt_Hausarzt')", HAECn, adOpenStatic, adLockReadOnly
'      rDMP.Open "SELECT genehmigung FROM " & hadbname & ".bs LEFT JOIN " & hadbname & ".arzt_has_bs ahb2 ON bs.idbs = ahb2.bs_id LEFT JOIN " & hadbname & ".arzt a2 ON ahb2.arzt_id = a2.idarzt LEFT JOIN " & hadbname & ".arzt_has_genehmigung ahg ON ahg.arzt_id = a2.idarzt LEFT JOIN " & hadbname & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE LEFT(bs.bsnr,7) = '" & Left$(üwerg(0, runde), 7) & "' AND genehmigung IN ('DMP-DM1_Koordinierender Arzt(SPP)_Erwachsene','DMP-DM1_Koordinierender Arzt(SPP)_Kinder/Jugendliche','DMP-DM1_Koordinierender Arzt_Hausarzt')", HAECn, adOpenStatic, adLockReadOnly
'      InfRoh(7, runde) = IIF(rDMP.BOF, vNS, "X")
'      SET rDMP = Nothing
'
'      gefunden = 0
'      HACngef = 0
'      For irunde = -1 To 1 ' aktuelle oder alte Hausärzte
'       IF Not gefunden AND üwerg(1, runde) <> vNS AND üwerg(2, runde) <> vNS THEN ' Nachname, Vorname
'        SET rHa = Nothing
'        IF irunde = -1 THEN
'         rHa.Open "SELECT CAST((SELECT GROUP_CONCAT(genehmigung,',') FROM " & hadbname & ".arzt_has_genehmigung ahg LEFT JOIN " & hadbname & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE ahg.arzt_id = a.idarzt) AS char) beme, GROUP_CONCAT(a1.nachname) gemmit, IF(nlart.niederlassungsart LIKE 'Fach%','FA','HA') arzttyp, ort.ort,`bs`.`BSNR` KVNu,REPLACE(`tel1`.`Tel`,'-','') AS `tel1`,REPLACE(`fax1`.`Fax`,'-','') AS `fax1`, mail.mail email, IF(`a`.`obweibl`,'Frau','Herr') AS `anrede`,`t`.`Titel` AS `titel`,`a`.`Vorname` AS `vorname`,`a`.`Nachname` AS `nachname`,CONCAT(`bs`.`Straße`,' ',`bs`.`Hausnr`) AS `straße`,`bs`.`PLZ` AS `plz`,GROUP_CONCAT(DISTINCT `fr`.`Fachrichtung` SEPARATOR ',') AS `zulg`,`bs`.`aktzeit` AS `aktzeit` " & _
'                  "FROM " & hadbname & ".bs LEFT JOIN " & hadbname & ".arzt_has_bs ahb ON ahb.bs_id = bs.idbs LEFT JOIN " & hadbname & ".arzt a ON a.idarzt = ahb.arzt_id LEFT JOIN " & hadbname & ".`ort` ON `bs`.`Ort_id` = `ort`.`idOrt` LEFT JOIN " & hadbname & ".`tel` `tel1` ON `tel1`.`bs_id` = `bs`.`idbs` LEFT JOIN " & hadbname & ".`fax` `fax1` ON `fax1`.`bs_id` = `bs`.`idbs` LEFT JOIN " & hadbname & ".mail mail ON mail.bs_id = bs.idbs LEFT JOIN " & hadbname & ".`titel` `t` ON `t`.`idtitel` = `a`.`titel_id` & _
                   "LEFT JOIN " & hadbname & ".`arzt_has_fachrichtung` `ahf` ON `ahf`.`arzt_id` = `a`.`idarzt` LEFT JOIN " & hadbname & ".`fachrichtung` `fr` ON `fr`.`idFachrichtung` = `ahf`.`fachrichtung_id` LEFT JOIN " & hadbname & ".nlart ON a.nlart_id = nlart.idnlart LEFT JOIN " & hadbname & ".arzt_has_bs ahb1 ON ahb1.bs_id = bs.idbs LEFT JOIN " & hadbname & ".arzt a1 ON a1.idarzt = ahb1.arzt_id AND a1.idarzt <> ahb.arzt_id WHERE LEFT(bsnr,7) = " & Left$(üwerg(0, runde), 7) & " AND a.nachname = '" & üwerg(1, runde) & "' AND a.vorname = '" & üwerg(2, runde) & "'", _
'                  HAECn, adOpenStatic, adLockReadOnly
'        Else
'         rHa.Open "SELECT h.* FROM hae" & IIF(irunde = 1, "alt", vNS) & " h WHERE kvnr = '" & LEFT(üwerg(0, runde), 2) & "/" & Right$(üwerg(0, runde), 5) & "' AND nachname = '" & üwerg(1, runde) & "' AND vorname = '" & üwerg(2, runde) & "'", HAECn, adOpenStatic, adLockReadOnly '  -dmpt1 AS j_dmpt1, -dmpt2 AS j_dmpt2,
'        END IF
'        IF Not rHa.BOF THEN IF NOT ISNULL(rHa!kvnu) THEN gefunden = True: Exit For
'       END IF
'       IF Not gefunden AND üwerg(1, runde) <> vNS THEN
'        SET rHa = Nothing
'        IF irunde = -1 THEN
'         rHa.Open "SELECT CAST((SELECT GROUP_CONCAT(genehmigung,',') FROM " & hadbname & ".arzt_has_genehmigung ahg LEFT JOIN " & hadbname & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE ahg.arzt_id = a.idarzt) AS char) beme, GROUP_CONCAT(a1.nachname) gemmit, IF(nlart.niederlassungsart LIKE 'Fach%','FA','HA') arzttyp, ort.ort,`bs`.`BSNR` KVNu,REPLACE(`tel1`.`Tel`,'-','') AS `tel1`,REPLACE(`fax1`.`Fax`,'-','') AS `fax1`, mail.mail email, IF(`a`.`obweibl`,'Frau','Herr') AS `anrede`,`t`.`Titel` AS `titel`,`a`.`Vorname` AS `vorname`,`a`.`Nachname` AS `nachname`,CONCAT(`bs`.`Straße`,' ',`bs`.`Hausnr`) AS `straße`,`bs`.`PLZ` AS `plz`,GROUP_CONCAT(DISTINCT `fr`.`Fachrichtung` SEPARATOR ',') AS `zulg`,`bs`.`aktzeit` AS `aktzeit` " & _
'                  "FROM " & hadbname & ".bs LEFT JOIN " & hadbname & ".arzt_has_bs ahb ON ahb.bs_id = bs.idbs LEFT JOIN " & hadbname & ".arzt a ON a.idarzt = ahb.arzt_id LEFT JOIN " & hadbname & ".`ort` ON `bs`.`Ort_id` = `ort`.`idOrt` LEFT JOIN " & hadbname & ".`tel` `tel1` ON `tel1`.`bs_id` = `bs`.`idbs` LEFT JOIN " & hadbname & ".`fax` `fax1` ON `fax1`.`bs_id` = `bs`.`idbs` LEFT JOIN " & hadbname & ".mail mail ON mail.bs_id = bs.idbs LEFT JOIN " & hadbname & ".`titel` `t` ON `t`.`idtitel` = `a`.`titel_id` LEFT JOIN " & hadbname & ".`arzt_has_fachrichtung` `ahf` ON `ahf`.`arzt_id` = `a`.`idarzt` LEFT JOIN " & hadbname & ".`fachrichtung` `fr` ON `fr`.`idFachrichtung` = `ahf`.`fachrichtung_id`  LEFT JOIN " & hadbname & ".nlart ON a.nlart_id = nlart.idnlart LEFT JOIN " & hadbname & ".arzt_has_bs ahb1 ON ahb1.bs_id = bs.idbs " & _
'                   "LEFT JOIN " & hadbname & ".arzt a1 ON a1.idarzt = ahb1.arzt_id AND a1.idarzt <> ahb.arzt_id WHERE LEFT(bsnr,7) = " & Left$(üwerg(0, runde), 7) & " AND a.nachname = '" & üwerg(1, runde) & "'", _
'                  HAECn, adOpenStatic, adLockReadOnly
'        Else
'         rHa.Open "SELECT h.* FROM hae" & IIF(irunde = 1, "alt", vNS) & " h WHERE kvnr = '" & LEFT(üwerg(0, runde), 2) & "/" & Right$(üwerg(0, runde), 5) & "' AND nachname = '" & üwerg(1, runde) & "'", HAECn, adOpenStatic, adLockReadOnly ' -dmpt1 AS j_dmpt1, -dmpt2 AS j_dmpt2,
'        END IF
'        IF Not rHa.BOF THEN IF NOT ISNULL(rHa!kvnu) THEN gefunden = True: Exit For
'       END IF
'       IF Not gefunden THEN
'        SET rHa = Nothing
'        IF irunde = -1 THEN
'         rHa.Open "SELECT CAST((SELECT GROUP_CONCAT(genehmigung,',') FROM " & hadbname & ".arzt_has_genehmigung ahg LEFT JOIN " & hadbname & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE ahg.arzt_id = a.idarzt) AS char) beme, GROUP_CONCAT(a1.nachname) gemmit, IF(nlart.niederlassungsart LIKE 'Fach%','FA','HA') arzttyp, ort.ort,`bs`.`BSNR` KVNu,REPLACE(`tel1`.`Tel`,'-','') AS `tel1`,REPLACE(`fax1`.`Fax`,'-','') AS `fax1`, mail.mail email, IF(`a`.`obweibl`,'Frau','Herr') AS `anrede`,`t`.`Titel` AS `titel`,`a`.`Vorname` AS `vorname`,`a`.`Nachname` AS `nachname`,CONCAT(`bs`.`Straße`,' ',`bs`.`Hausnr`) AS `straße`,`bs`.`PLZ` AS `plz`,GROUP_CONCAT(DISTINCT `fr`.`Fachrichtung` SEPARATOR ',') AS `zulg`,`bs`.`aktzeit` AS `aktzeit` " & _
'                  "FROM " & hadbname & ".bs LEFT JOIN " & hadbname & ".arzt_has_bs ahb ON ahb.bs_id = bs.idbs LEFT JOIN " & hadbname & ".arzt a ON a.idarzt = ahb.arzt_id LEFT JOIN " & hadbname & ".`ort` ON `bs`.`Ort_id` = `ort`.`idOrt` LEFT JOIN " & hadbname & ".`tel` `tel1` ON `tel1`.`bs_id` = `bs`.`idbs` LEFT JOIN " & hadbname & ".`fax` `fax1` ON `fax1`.`bs_id` = `bs`.`idbs` LEFT JOIN " & hadbname & ".mail mail ON mail.bs_id = bs.idbs LEFT JOIN " & hadbname & ".`titel` `t` ON `t`.`idtitel` = `a`.`titel_id` LEFT JOIN " & hadbname & ".`arzt_has_fachrichtung` `ahf` ON `ahf`.`arzt_id` = `a`.`idarzt` LEFT JOIN " & hadbname & ".`fachrichtung` `fr` ON `fr`.`idFachrichtung` = `ahf`.`fachrichtung_id`  LEFT JOIN " & hadbname & ".nlart ON a.nlart_id = nlart.idnlart LEFT JOIN " & hadbname & ".arzt_has_bs ahb1 ON ahb1.bs_id = bs.idbs LEFT JOIN " & hadbname & ".arzt a1 ON a1.idarzt = ahb1.arzt_id AND a1.idarzt <> ahb.arzt_id WHERE LEFT(bsnr,7) = " & Left$(üwerg(0, runde), 7), _
'                  HAECn, adOpenStatic, adLockReadOnly
'        Else
'         rHa.Open "SELECT h.* FROM hae" & IIF(irunde = 1, "alt", vNS) & " h WHERE kvnr = '" & LEFT(üwerg(0, runde), 2) & "/" & Right$(üwerg(0, runde), 5) & "'", HAECn, adOpenStatic, adLockReadOnly ' -dmpt1 AS j_dmpt1, -dmpt2 AS j_dmpt2,
'        END IF
'        IF Not rHa.BOF THEN IF NOT ISNULL(rHa!kvnu) THEN gefunden = True: Exit For
'       END IF
'      Next irunde
'      IF gefunden THEN
'       IF NOT ISNULL(rHa!anrede) THEN
'        InfRoh(0, runde) = rHa!anrede
'       END IF
'       InfRoh(1, runde) = IIF(ISNULL(rHa!titel) OR LenB(rHa!titel) = 0, "Dr.med.", rHa!titel) & " " & rHa!vorname & " " & rHa!nachname
'       IF NOT ISNULL(rHa!straße) THEN
'        InfRoh(2, runde) = rHa!straße
'       END IF
'       InfRoh(3, runde) = rHa!plz & " " & rHa!ort
'       IF InfRoh(4, runde) = vNS AND NOT ISNULL(rHa!fax1) THEN
'        InfRoh(4, runde) = rHa!fax1
'       END IF
''       InfRoh(6, runde) = IIF(rHa!j_dmpt2 <> 0, "X", vNS) ' geht auch für null
''       InfRoh(7, runde) = IIF(rHa!j_dmpt1 <> 0, "X", vNS)
''       SET rDMP = Nothing
''       SET rDMP = New ADODB.Recordset
''       rDMP.Open "SELECT * FROM hae WHERE kvnr = '" & LEFT(üwerg(0, runde), 2) & "/" & Right$(üwerg(0, runde), 5) & "' AND dmpt2 = 1", HAECn, adOpenStatic, adLockReadOnly
''       InfRoh(6, runde) = IIF(rDMP.BOF, vNS, "X")
''       SET rDMP = Nothing
''       SET rDMP = New ADODB.Recordset
''       rDMP.Open "SELECT * FROM hae WHERE kvnr = '" & LEFT(üwerg(0, runde), 2) & "/" & Right$(üwerg(0, runde), 5) & "' AND dmpt1 = 1", HAECn, adOpenStatic, adLockReadOnly
''       InfRoh(7, runde) = IIF(rDMP.BOF, vNS, "X")
'       IF LenB(InfRoh(8, runde)) = 0 THEN InfRoh(8, runde) = IIF(ISNULL(rHa!zulg), vNS, rHa!zulg)
'       IF LenB(InfRoh(9, runde)) = 0 THEN InfRoh(9, runde) = rHa!vorname
'       IF LenB(InfRoh(14, runde)) = 0 THEN InfRoh(14, runde) = rHa!nachname
'       InfRoh(10, runde) = üwerg(3, runde)
'       InfRoh(12, runde) = Left$(rHa!kvnu, 7)
'       IF NOT ISNULL(rHa!tel1) THEN
'        InfRoh(13, runde) = rHa!tel1
'       END IF
'       HACngef = -1
'      END IF
'
'      gefunden = 0
'      IF üwerg(1, runde) <> vNS AND üwerg(2, runde) <> vNS THEN
'       SET rHae = Nothing
'       myFrag rHae, "SELECT * FROM `hausaerzte` WHERE LEFT(kvnr,7) = '" & Left$(üwerg(0, runde), 7) & "' AND nachname = '" & üwerg(1, runde) & "' AND vorname = '" & üwerg(2, runde) & "'"
'       IF Not rHae.BOF THEN gefunden = True
'      END IF
'      IF Not gefunden AND üwerg(1, runde) <> vNS THEN
'       SET rHae = Nothing
'       myFrag rHae, "SELECT * FROM `hausaerzte` WHERE LEFT(kvnr,7) = '" & Left$(üwerg(0, runde), 7) & "' AND nachname = '" & üwerg(1, runde) & "'"
'       IF Not rHae.BOF THEN gefunden = True
'      END IF
'      IF Not gefunden THEN
'       SET rHae = Nothing
'       myFrag rHae, "SELECT * FROM `hausaerzte` WHERE LEFT(kvnr,7) = '" & Left$(üwerg(0, runde), 7) & "'"
'       IF Not rHae.BOF THEN gefunden = True
'      END IF
'      IF gefunden THEN
'       IF InfRoh(0, runde) = vNS THEN InfRoh(0, runde) = IIF(rHae!Geschlecht = "w", "Frau", "Herr")
'       IF InfRoh(1, runde) = vNS THEN InfRoh(1, runde) = IIF(ISNULL(rHae!titel) OR LenB(rHae!titel) = 0, "Dr.med.", rHae!titel) & " " & rHae!vorname & " " & rHae!nachname
'       IF InfRoh(2, runde) = vNS THEN InfRoh(2, runde) = IIF(ISNULL(rHae!straße) OR LenB(rHae!straße) = 0, vNS, rHae!straße)
'       IF InfRoh(3, runde) = vNS THEN InfRoh(3, runde) = rHae!plz & " " & rHae!ort
'       IF InfRoh(9, runde) = vNS THEN InfRoh(9, runde) = IIF(ISNULL(rHae!vorname) OR LenB(rHae!vorname) = 0, vNS, rHae!vorname)
'       IF InfRoh(14, runde) = vNS THEN InfRoh(14, runde) = IIF(ISNULL(rHae!Name) OR LenB(rHae!Name) = 0, vNS, rHae!Name)
'       IF InfRoh(4, runde) = vNS AND NOT ISNULL(rHae!telefax) THEN
'        InfRoh(4, runde) = rHae!telefax
'       END IF
'        SELECT CASE rHae!Überschrift
'         Case "L": InfRoh(5, runde) = IIF(InfRoh(0, runde) = "Frau", "Liebe", "Lieber") & " " & InfRoh(9, runde) ' nicht: rHae!Geschlecht = "w"
'         Case "H": InfRoh(5, runde) = "Hallo " + InfRoh(9, runde)
'         Case Else: InfRoh(5, runde) = IIF(InfRoh(0, runde) = "Frau", "Sehr geehrte Frau Kollegin", "Sehr geehrter Herr Kollege")
'        END SELECT
''       infroh(6, runde) = IIF(rHae!dmpt2 = -1, "X", vns)
''       infroh(7, runde) = IIF(rHae!dmpt1 = -1, "X", vns)
'       IF InfRoh(8, runde) = vNS THEN InfRoh(8, runde) = IIF(ISNULL(rHae!zulassungsgebiet), vNS, rHae!zulassungsgebiet)
'       InfRoh(10, runde) = üwerg(3, runde)
'       InfRoh(12, runde) = rHae!KVNr
'       InfRoh(13, runde) = IIF(ISNULL(rHae!telefon), vNS, rHae!telefon)
'      END IF
'      Dim obInsert%
'      IF Not gefunden THEN
'       obInsert = True
'      Else
'       obInsert = False
'       IF Not rHae.BOF THEN
'        IF rHae!nachname <> üwerg(1, runde) OR rHae!vorname <> üwerg(2, runde) THEN obInsert = True
'       END IF
'      END IF
'      IF obInsert AND NOT rHae.BOF THEN
'       IF HACngef THEN
'        SET rHae = Nothing
''        Call myEFrag("SET foreign_key_checks = 0")
''        Call myEFrag("DELETE FROM `hausaerzte` WHERE id >= 753")
''        Call myEFrag("SET foreign_key_checks = 1")
'        InsKorr DBCn, "INSERT INTO `hausaerzte`(name, vorname, nachname, anschrift, kvnr, telefon, telefax, e_mail, zulassungsgebiet, arzttyp, " & _
'        "`gemeinschaftspraxis mit`" & ", beme, geschlecht,titel,straße,ort,plz,überschrift,dmpt2,dmpt1,zahl,nichtmehr,schwerpunkt,zusatzbezeichnung,bemerkung,sprechstunden) VALUES('" & rHa!anrede & " " & IIF(rHa!titel <> vNS AND NOT ISNULL(rHa!titel), rHa!titel, "Dr.med.") & " " & rHa!vorname & " " & rHa!nachname & "','" & rHa!vorname & "','" & rHa!nachname & "','" & rHa!straße & ", " & rHa!plz & " " & rHa!ort & "','" & rHa!kvnu & "','" & rHa!tel1 & "','" & rHa!fax1 & "','" & rHa!email & "','" & rHa!zulg & "','" & rHa!arzttyp & "','" & rHa!gemmit & "','" & rHa!beme & "','" & IIF(rHa!anrede = "w", "Frau", "Herr") & "','" & rHa!titel & "','" & rHa!straße & "','" & rHa!ort & "','" & rHa!plz & "',''," & IIF(InfRoh(6, runde) = "X", "1", "0") & "," & IIF(InfRoh(7, runde) = "X", "1", "0") & ",0,0,'','','','')", rAF
'        obInsert = False
'       END IF
'      END IF
'     END IF ' üwerg(0,runde)<> vns
'     IF InfRoh(5, runde) = vNS THEN InfRoh(5, runde) = IIF(InfRoh(0, runde) = "Frau", "Sehr geehrte Frau Kollegin", IIF(InfRoh(0, runde) = "Herr", "Sehr geehrter Herr Kollege", "Sehr geehrte Frau Kollegin, sehr geehrter Herr Kollege"))
'     InfRoh(11, runde) = IIF(InStrB(LCase$(InfRoh(8, runde)), "intern") <> 0 OR InStrB(LCase$(InfRoh(8, runde)), "innere") <> 0 OR InStrB(LCase$(InfRoh(8, runde)), "allg") <> 0 OR InStrB(LCase$(InfRoh(8, runde)), "prakt") <> 0 OR InfRoh(8, runde) = "Arzt" OR InfRoh(8, runde) = "Ärztin" OR InfRoh(8, runde) = "Hausarzt" OR InfRoh(8, runde) = "Hausärztin", "X", vNS)
'    Next runde
'    Dim pos&, i%, aktrunde%, gelöscht%
'    pos = UBound(InfRoh, 2)
'' zuletzt kommen die Überweiser, die nicht als Allgemeinarzt/Internist sind und nicht als Hausarzt eingetragen
'    For runde = UBound(InfRoh, 2) To 0 Step -1
'     IF InfRoh(11, runde) = vNS AND InStrB(InfRoh(10, runde), "HA") = 0 THEN
'      For i = 0 To UBound(InfRoh, 1)
'       Infos(i, pos) = InfRoh(i, runde)
''       InfRoh(i, runde) = vns
'      Next i
'      pos = pos - 1
'     END IF
'    Next runde
'    aktrunde = UBound(InfRoh, 2) + 1 ' der letzte hier berücksichtigte
'' dann kommen bis zur 3. Stelle die Überweiser, die als Allgemeinarzt/Internist sind aber nicht als Hausarzt eingetragen
'    For runde = UBound(InfRoh, 2) To 0 Step -1
'     IF pos <= IIF(obHAPrio, 0, 1) THEN Exit For
'     IF InfRoh(11, runde) = "X" AND InStrB(InfRoh(10, runde), "HA") = 0 THEN
'      gelöscht = 0
'      For irunde = UBound(InfRoh, 2) To 0 Step -1
'       IF irunde <> runde THEN
'        IF InfRoh(1, irunde) = InfRoh(1, runde) THEN ' doppelter Arzt
'         For i = 0 To UBound(InfRoh, 1)
'          InfRoh(i, runde) = vNS
'         Next i
'         gelöscht = True
'         Exit For
'        END IF
'       END IF
'      Next irunde
'      IF Not gelöscht THEN ' war ja oben so, also falsch nicht gelöscht
'       For i = 0 To UBound(InfRoh, 1)
'        Infos(i, pos) = InfRoh(i, runde)
'        InfRoh(i, runde) = vNS ' schon eingetragen
'       Next i
'      END IF
'      pos = pos - 1
'     END IF
'     aktrunde = runde
'    Next runde
'' dann kommt der Hausarzt, falls bisher nicht berücksichtigt
'    For runde = UBound(InfRoh, 2) To aktrunde Step -1
'     IF InStrB(InfRoh(10, runde), "HA") <> 0 THEN
'      gelöscht = 0
'      For irunde = UBound(InfRoh, 2) To 0 Step -1
'       IF irunde <> runde THEN
'        IF InfRoh(1, irunde) = InfRoh(1, runde) AND NOT (InStrB(InfRoh(10, irunde), "HA") = 0 AND InStrB(InfRoh(10, runde), "HA") <> 0) THEN ' doppelter Arzt, Einschränkung 8.2.10
'         For i = 0 To UBound(InfRoh, 1)
'          IF i = 10 THEN
'           InfRoh(i, irunde) = InfRoh(i, irunde) & "," & InfRoh(i, runde)
'          END IF
'          InfRoh(i, runde) = vNS
'         Next i
'         gelöscht = True
'         Exit For
'        END IF
'       END IF
'      Next irunde
'      IF Not gelöscht THEN ' war ja oben so, also falsch nicht gelöscht
'       For i = 0 To UBound(InfRoh, 1)
'        Infos(i, pos) = InfRoh(i, runde)
'        InfRoh(i, runde) = vNS
'       Next i
'      END IF
'      pos = pos - 1
'      Exit For
'     END IF
'    Next runde
'' dann der Rest in der Reihenfolge
'    For runde = aktrunde - 1 To 0 Step -1
'     IF InfRoh(11, runde) = "X" OR InStrB(InfRoh(10, runde), "HA") <> 0 THEN
'      gelöscht = 0
'      For irunde = UBound(InfRoh, 2) To 0 Step -1
'       IF irunde <> runde THEN
'        IF InfRoh(1, irunde) = InfRoh(1, runde) AND NOT (InStrB(InfRoh(10, irunde), "HA") = 0 AND InStrB(InfRoh(10, runde), "HA") <> 0) THEN ' doppelter Arzt, Einschränkung 28.8.08
'         For i = 0 To UBound(InfRoh, 1)
'          IF i = 10 THEN
'           InfRoh(i, irunde) = InfRoh(i, irunde) & "," & InfRoh(i, runde)
'          END IF
'          InfRoh(i, runde) = vNS
'         Next i
'         gelöscht = True
'         Exit For
'        END IF
'       END IF
'      Next irunde
'      IF Not gelöscht THEN  ' war ja oben so, also falsch nicht gelöscht
'       For i = 0 To UBound(InfRoh, 1)
'        Infos(i, pos) = InfRoh(i, runde)
'        InfRoh(i, runde) = vNS
'       Next i
'      END IF
'      pos = pos - 1
'     END IF
'    Next runde
'' Wenn erste Position leer, schaut's unschön aus
'    IF Infos(1, 0) = vNS THEN
'     IF UBound(Infos, 2) > 0 THEN
'      IF Infos(1, 1) <> vNS THEN
'       For i = 0 To UBound(Infos, 1)
'        Infos(i, 0) = Infos(i, 1)
'        Infos(i, 1) = vNS
'       Next i
'      Else
'       IF UBound(Infos, 2) > 1 THEN
'        IF Infos(1, 2) <> vNS THEN
'         For i = 0 To UBound(Infos, 1)
'          Infos(i, 0) = Infos(i, 1)
'          Infos(i, 1) = vNS
'         Next i
'        END IF
'       END IF
'      END IF
'     END IF
'    END IF
'    Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIF(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getHausarztAlt/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' getHausarztAlt

'Aufruf in: alleSpeichern, dodoPlz, harealneu falschebriefelöschen, doVerdächtigeÜberweiser,  gethatest
Function getHausarzt1(infos$(), rFa() As Faelle, rKv() As kvnrue, Optional obHAPrio%, Optional Pat_ID, Optional auchwir% = 0, Optional QZahl% = 0)
' Dim rNa As New ADODB.Recordset
 Dim rHa As New ADODB.Recordset, rAf&
 Dim rsFa As New ADODB.Recordset
 Dim gefunden%, runde%, irunde%, HACngef%
 Dim rListena As New ADODB.Recordset
 Dim rhae As New ADODB.Recordset
 Dim faxfehlt%
 Dim hains() As hatyp
' Dim rFa As New ADODB.Recordset
' Dim i%, j%
 Dim InfRoh$()
 Dim übwerg$() ' HaErg$()
' Dim Tj!(120), m%, p%
 ' m = 0: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
 On Error GoTo fehler
 If Not IsMissing(Pat_ID) Then
  ReDim rFa(0)
  ' m = 1: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
  myFrag rsFa, "SELECT * FROM (SELECT * FROM faelle WHERE pat_id=" & Pat_ID & " ORDER BY bhfb DESC " & IIf(QZahl, " LIMIT " & QZahl, "") & ") i ORDER BY bhfb"
  ' m = 2: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
  If Not rsFa.BOF Then
   Do While Not rsFa.EOF
'   If rNa(0).Pat_id = 68076 Then Stop
    ReDim Preserve rFa(UBound(rFa) + 1)
    rFa(UBound(rFa)).Pat_ID = rsFa!Pat_ID
    rFa(UBound(rFa)).BhFB = rsFa!BhFB
    rFa(UBound(rFa)).AndÜw = rsFa!AndÜw
    rFa(UBound(rFa)).Übwr = rsFa!Übwr         ' IF(übwvbsnr='',IF(übwvkvnr = '', andüw, übwvkvnr),übwvbsnr) üw
    rFa(UBound(rFa)).ÜbWVBSNR = rsFa!ÜbWVBSNR
    rFa(UBound(rFa)).ÜbWVLANR = rsFa!ÜbWVLANR
    rFa(UBound(rFa)).ÜbWVKVNR = rsFa!ÜbWVKVNR
    rFa(UBound(rFa)).ÜWNaN = rsFa!ÜWNaN
    rFa(UBound(rFa)).ÜWVor = rsFa!ÜWVor
    rFa(UBound(rFa)).Quartal = rsFa!Quartal
    rFa(UBound(rFa)).ÜWVor = rsFa!ÜWVor
    rsFa.MoveNext
   Loop
  End If
 End If
 ' m = 3: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
 ReDim hains(0)
 ReDim infos(15, 2)
 '  0: Frau/Herrn
 '  1: Titel+Vorn+Nachn,
 '  2: Straße,
 '  3: PLZ+Ort,
 '  4: Faxnr,
 '  5: S.g./Liebe,
 '  6: DMPTyp2,
 '  7: DMPTyp1,
 '  8: Niederlassungsgebiet 3. Feld für einmal austauschen,
 '  9: Vorname,
 ' 10: Funktion ("Üw 207, HA"),
 ' 11: Fachrichtung
 ' 12: KV-Nummer
 ' 13: Tel'nr.
 ' 14: Nachname,
 ' 15: Email
 

' in Faelle steht der als Überweiser eingetragene
 If Üwrd(rFa(), übwerg, rKv, auchwir) = vNS Then Exit Function
 For irunde = 0 To UBound(übwerg, 2)
  übwerg(1, irunde) = Trim$(übwerg(1, irunde))
  übwerg(2, irunde) = Trim$(übwerg(2, irunde))
  übwerg(3, irunde) = Trim$(übwerg(3, irunde))
  übwerg(4, irunde) = Trim$(übwerg(4, irunde))
 Next irunde
 ' m = 4: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
    
    ReDim InfRoh(15, UBound(übwerg, 2))
    ReDim infos(15, UBound(übwerg, 2))
    For runde = 0 To UBound(übwerg, 2)
 ' m = 5 + 10 * runde: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
     If übwerg(0, runde) <> vNS Then
     
      gefunden = 0
      If übwerg(1, runde) <> vNS And übwerg(2, runde) <> vNS Then
       Set rListena = Nothing
'       myFrag rListena, "SELECT * FROM `liuez` WHERE LEFT(kvnr,7) = '" & Left$(übwerg(0, runde), 7) & "' AND name = '" & übwerg(1, runde) & "' AND vorname = '" & übwerg(2, runde) & "' ORDER BY fax DESC"
       myFrag rListena, "SELECT * FROM `aktlue` WHERE kvnro = '" & übwerg(4, runde) & "' AND nameo = '" & übwerg(1, runde) & "' AND vno = '" & übwerg(2, runde) & "' ORDER BY fax DESC"
       If Not rListena.BOF Then gefunden = True
      End If
      If Not gefunden And übwerg(1, runde) <> vNS Then
       Set rListena = Nothing
'       myFrag rListena, "SELECT * FROM `liuez` WHERE LEFT(kvnr,7) = '" & Left$(übwerg(0, runde), 7) & "' AND name = '" & übwerg(1, runde) & "' ORDER BY fax DESC"
       myFrag rListena, "SELECT * FROM `aktlue` WHERE kvnro = '" & übwerg(4, runde) & "' AND nameo = '" & übwerg(1, runde) & "' ORDER BY fax DESC"
       If Not rListena.BOF Then gefunden = True
      End If
      If Not gefunden And übwerg(4, runde) <> vNS Then
       Set rListena = Nothing
'       myFrag rListena, "SELECT * FROM `liuez` WHERE LEFT(kvnr,7) = '" & Left$(übwerg(0, runde), 7) & "' ORDER BY fax DESC"
       myFrag rListena, "SELECT * FROM `aktlue` WHERE kvnro = '" & übwerg(4, runde) & "' ORDER BY fax DESC"
       If Not rListena.BOF Then gefunden = True
      End If
      If Not gefunden And übwerg(1, runde) <> vNS Then
       Set rListena = Nothing
       myFrag rListena, "SELECT * FROM `aktlue` WHERE nameo = '" & übwerg(1, runde) & "' AND vno = '" & übwerg(2, runde) & "' ORDER BY fax DESC"
       If Not rListena.BOF Then gefunden = True
      End If
 ' m = 6 + 10 * runde: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
      If Not gefunden Then
       Set rListena = Nothing
       Dim sql$
       sql = "SELECT titel, Niederlassungsart arzttyp, GROUP_CONCAT(DISTINCT a2.nachname) gemmit, fr.fachrichtung fachgruppe, mail email, LEFT(bsnr,7) kvnr, fax, REPLACE(GROUP_CONCAT(DISTINCT tel),'-','') tel, REPLACE(GROUP_CONCAT(DISTINCT fax),'-','') fax, a.nachname name, a.vorname, bs.straße strasse, bs.plz, ort, IF(a.obweibl=1,'w','m') geschlecht, IF(a.obweibl=1,'Frau','Herrn') anrede, tel telefon, GROUP_CONCAT(DISTINCT fachrichtung) zulg, " & _
           "MAX(IF(ISNULL((SELECT genehmigung FROM " & HADBName & ".arzt_has_genehmigung ahg1 LEFT JOIN " & HADBName & ".genehmigung g1 ON g1.idgenehmigung = ahg1.genehmigung_id WHERE ahg1.arzt_id = a2.idarzt AND genehmigung = 'DMP-DM1_Koordinierender Arzt_Hausarzt')),'0','1')) dmpt1, " & _
           "MAX(IF(ISNULL((SELECT genehmigung FROM " & HADBName & ".arzt_has_genehmigung ahg2 LEFT JOIN " & HADBName & ".genehmigung g2 ON g2.idgenehmigung = ahg2.genehmigung_id WHERE ahg2.arzt_id = a2.idarzt AND genehmigung = 'DMP-DM2_Koordinierender Arzt')),'0','1')) dmpt2 " & _
           "FROM " & HADBName & ".arzt a " & _
           "LEFT JOIN " & HADBName & ".arzt_has_fachrichtung ahf ON a.idarzt = ahf.arzt_id LEFT JOIN " & HADBName & ".fachrichtung fr ON ahf.fachrichtung_id = fr.idfachrichtung " & _
           "LEFT JOIN " & HADBName & ".arzt_has_bs ahb ON a.idarzt = ahb.arzt_id LEFT JOIN " & HADBName & ".bs ON bs.idbs = ahb.bs_id " & _
           "LEFT JOIN " & HADBName & ".tel ON tel.bs_id = bs.idbs " & _
           "LEFT JOIN " & HADBName & ".fax ON fax.bs_id = bs.idbs " & _
           "LEFT JOIN " & HADBName & ".mail ON mail.bs_id = bs.idbs " & _
           "LEFT JOIN " & HADBName & ".arzt_has_bs ahb2 ON bs.idbs = ahb2.bs_id LEFT JOIN " & HADBName & ".arzt a2 ON a2.idarzt = ahb2.arzt_id " & _
           "LEFT JOIN " & HADBName & ".nlart ON a.nlart_id = nlart.idnlart " & _
           "LEFT JOIN " & HADBName & ".titel t ON t.idtitel = a.titel_id " & _
           "LEFT JOIN " & HADBName & ".ort o ON o.idort = bs.ort_id " & _
           "WHERE LEFT(bsnr,7) = '" & Left$(übwerg(0, runde), 7) & "'"
       myFrag rListena, sql
       If Not rListena.BOF Then gefunden = True
      End If
  ' m = 7 + 10 * runde: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
     If gefunden Then
       InfRoh(1, runde) = IIf(IsNull(rListena!Titel) Or LenB(rListena!Titel) = 0, "Dr.med.", rListena!Titel) & " " & rListena!Vorname & " " & rListena!name
       InfRoh(0, runde) = IIf(IsNull(rListena!anrede), "", rListena!anrede)
       If Not IsNull(rListena!strasse) Then
        InfRoh(2, runde) = rListena!strasse
        InfRoh(3, runde) = rListena!plz & " " & rListena!ort
        If IsNull(rListena!fax) Then
         Dim rfax As New ADODB.Recordset
         Set rfax = Nothing
         myFrag rfax, "SELECT faxzahl FROM " & HADBName & ".bs bsr LEFT JOIN " & HADBName & ".arzt_has_bs ahbr ON bsr.idbs = ahbr.bs_id LEFT JOIN " & HADBName & ".arzt a ON a.idarzt = ahbr.arzt_id LEFT JOIN " & HADBName & ".arzt_has_bs ahb ON a.idarzt = ahb.arzt_id LEFT JOIN " & HADBName & ".bs ON bs.idbs = ahb.bs_id LEFT JOIN " & HADBName & ".fax ON fax.bs_id = bsr.idbs WHERE LEFT(bs.bsnr,7) = '" & Left$(übwerg(0, runde), 7) & "'  AND a.obehem=0 AND NOT ISNULL(faxzahl)"
         If Not rfax.BOF Then
          InfRoh(4, runde) = rfax!faxzahl
         End If
        Else
         InfRoh(4, runde) = rListena!fax
        End If
       End If ' NOT ISNULL(rListena!strasse) Then
       ' folgendes am 25.4.22 aus NOT ISNULL(rlistena!strasse) ausgeglieder:
' folgende Unterscheidung eingefügt wg. F.Dr. Junger 5.12.11
        On Error Resume Next
        If rListena!KVNr = vNS Then
         If rFa(1).Übwr <> vNS Then
          InfRoh(12, runde) = rFa(1).Übwr
          myEFrag "UPDATE `liuez` SET kvnr = " & rFa(1).Übwr & " WHERE id = " & rListena!id, rAf
'          Debug.Print rAF
         End If
        Else ' rListena!KVNr = vNS Then
         InfRoh(12, runde) = rListena!KVNr
        End If ' rListena!KVNr = vNS Then Else
        On Error GoTo fehler
       If IsNull(rListena!fachgruppe) Then InfRoh(8, runde) = vNS Else InfRoh(8, runde) = rListena!fachgruppe
       If IsNull(rListena!Vorname) Then InfRoh(9, runde) = vNS Else InfRoh(9, runde) = rListena!Vorname
       If IsNull(rListena!name) Then InfRoh(14, runde) = vNS Else InfRoh(14, runde) = rListena!name
       InfRoh(10, runde) = übwerg(3, runde)
       If IsNull(rListena!telefon) Then
         Set rfax = Nothing
         myFrag rfax, "SELECT tel FROM " & HADBName & ".bs bsr LEFT JOIN " & HADBName & ".arzt_has_bs ahbr ON bsr.idbs = ahbr.bs_id LEFT JOIN " & HADBName & ".arzt a ON a.idarzt = ahbr.arzt_id LEFT JOIN " & HADBName & ".arzt_has_bs ahb ON a.idarzt = ahb.arzt_id LEFT JOIN " & HADBName & ".bs ON bs.idbs = ahb.bs_id LEFT JOIN " & HADBName & ".tel ON tel.bs_id = bsr.idbs WHERE (CONCAT(LEFT(bs.bsnr,7),'00') = '" & rFa(UBound(rFa)).ÜbWVBSNR & "' OR bs.bsnr='" & rFa(UBound(rFa)).ÜbWVBSNR & "') AND NOT ISNULL(tel) AND a.obehem=0"
         If Not rfax.BOF Then
          If Not IsNull(rListena!name) Then ' 4.7.17 kein Fax an NULL!
           InfRoh(4, runde) = rfax!tel
          End If ' NOT ISNULL(rListena!name) Then '
         End If ' Not rfax.BOF Then
       Else ' IsNull(rListena!telefon) Then
        InfRoh(13, runde) = rListena!telefon
       End If ' IsNull(rListena!telefon) Then else
      End If ' gefunden Then
  ' m = 8 + 10 * runde: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
      
      If VarType(übwerg(0, runde)) = vbString Then
       übwerg(0, runde) = CDbl(REPLACE$(REPLACE$(übwerg(0, runde), "/", ""), Space$(1), vNS))
      End If
'      If übwerg(0, runde) < 9999999 Then übwerg(0, runde) = übwerg(0, runde) * 100
      If übwerg(0, runde) > 6100000 And übwerg(0, runde) < 9000000 Then übwerg(0, runde) = übwerg(0, runde) * 100
      Dim rDMP As New ADODB.Recordset
      Set rDMP = Nothing
'      rDMP.Open "SELECT genehmigung FROM (" & hadbname & ".arzt a LEFT JOIN " & hadbname & ".arzt_has_bs ahb ON ahb.arzt_id = a.idarzt)  LEFT JOIN " & hadbname & ".bs ON ahb.bs_id = bs.idbs LEFT JOIN " & hadbname & ".arzt_has_bs ahb2 ON bs.idbs = ahb2.bs_id LEFT JOIN " & hadbname & ".arzt a2 ON ahb2.arzt_id = a2.idarzt LEFT JOIN " & hadbname & ".arzt_has_genehmigung ahg ON ahg.arzt_id = a2.idarzt LEFT JOIN " & hadbname & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE bs.bsnr = '" & left$(übwerg(0, runde),7) & "' AND genehmigung IN ('DMP-DM2_Koordinierender Arzt','DMP-DM2_FA_Diab.bes.qualif.Arzt(SPP)+Koordin. Arzt')", HAECn, adOpenStatic, adLockReadOnly
'      Call acon(HaT) ' z.B. für Genehmigung
      If LenB(DBCn) = 0 Or DBCn = "" Then Call acon(quelleT)
      myFrag rDMP, "SELECT genehmigung FROM " & HADBName & ".bs LEFT JOIN " & HADBName & ".arzt_has_bs ahb2 ON bs.idbs = ahb2.bs_id LEFT JOIN " & HADBName & ".arzt a2 ON ahb2.arzt_id = a2.idarzt LEFT JOIN " & HADBName & ".arzt_has_genehmigung ahg ON ahg.arzt_id = a2.idarzt LEFT JOIN " & HADBName & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE LEFT(bs.bsnr,7) = '" & Left$(übwerg(0, runde), 7) & "' AND genehmigung IN ('DMP-DM2_Koordinierender Arzt','DMP-DM2_FA_Diab.bes.qualif.Arzt(SPP)+Koordin. Arzt')  AND a2.obehem=0" 'haecn
      InfRoh(6, runde) = IIf(rDMP.BOF, vNS, "X")
      Set rDMP = Nothing
'      rDMP.Open "SELECT genehmigung FROM (" & hadbname & ".arzt a LEFT JOIN " & hadbname & ".arzt_has_bs ahb ON ahb.arzt_id = a.idarzt)  LEFT JOIN " & hadbname & ".bs ON ahb.bs_id = bs.idbs LEFT JOIN " & hadbname & ".arzt_has_bs ahb2 ON bs.idbs = ahb2.bs_id LEFT JOIN " & hadbname & ".arzt a2 ON ahb2.arzt_id = a2.idarzt LEFT JOIN " & hadbname & ".arzt_has_genehmigung ahg ON ahg.arzt_id = a2.idarzt LEFT JOIN " & hadbname & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE bs.bsnr = '" & left$(übwerg(0, runde),7) & "' AND genehmigung IN ('DMP-DM1_Koordinierender Arzt(SPP)_Erwachsene','DMP-DM1_Koordinierender Arzt(SPP)_Kinder/Jugendliche','DMP-DM1_Koordinierender Arzt_Hausarzt')", HAECn, adOpenStatic, adLockReadOnly
      myFrag rDMP, "SELECT genehmigung FROM " & HADBName & ".bs LEFT JOIN " & HADBName & ".arzt_has_bs ahb2 ON bs.idbs = ahb2.bs_id LEFT JOIN " & HADBName & ".arzt a2 ON ahb2.arzt_id = a2.idarzt LEFT JOIN " & HADBName & ".arzt_has_genehmigung ahg ON ahg.arzt_id = a2.idarzt LEFT JOIN " & HADBName & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE LEFT(bs.bsnr,7) = '" & Left$(übwerg(0, runde), 7) & "' AND genehmigung IN ('DMP-DM1_Koordinierender Arzt(SPP)_Erwachsene','DMP-DM1_Koordinierender Arzt(SPP)_Kinder/Jugendliche','DMP-DM1_Koordinierender Arzt_Hausarzt') AND a2.obehem=0" ' HAECn
      InfRoh(7, runde) = IIf(rDMP.BOF, vNS, "X")
      Set rDMP = Nothing
      
      HACngef = 0
  ' m = 9 + 10 * runde: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
      For irunde = -1 To 1 ' aktuelle oder alte Hausärzte
       If Not HACngef And übwerg(1, runde) <> vNS And übwerg(2, runde) <> vNS Then ' Nachname, Vorname
        Set rHa = Nothing
Const sql0$ = "SELECT " & _
                     "IF(nlart.niederlassungsart LIKE 'Fach%','FA','HA') `arzttyp`, " & _
                     "ort.ort, `bs`.`BSNR` `KVNu`, " & _
                     "REPLACE(`tel1`.`Tel`,'-','') `tel1`," & _
                     "REPLACE(`fax1`.`Fax`,'-','') `fax1`," & _
                     "mail.mail email, IF(`a`.`obweibl`,'Frau','Herr') `anrede`,`t`.`Titel` `titel`," & _
                     "`a`.`Vorname` AS `vorname`,`a`.`Nachname` `nachname`," & _
                     "CONCAT(`bs`.`Straße`,' ',`bs`.`Hausnr`) `straße`," & _
                     "`bs`.`PLZ` AS `plz`,GROUP_CONCAT(DISTINCT `fr`.`Fachrichtung` SEPARATOR ',') AS `zulg`," & _
                     "CAST((SELECT GROUP_CONCAT(genehmigung) FROM " & HADBName & ".arzt_has_genehmigung ahg LEFT JOIN " & HADBName & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE ahg.arzt_id = a.idarzt) AS char) `beme`," & _
                     "CAST((SELECT GROUP_CONCAT(nachname) FROM " & HADBName & ".arzt_has_bs ahb1 LEFT JOIN " & HADBName & ".arzt a1 ON a1.idarzt = ahb1.arzt_id WHERE ahb1.bs_id = bs.idbs AND a1.idarzt <> ahb.arzt_id) AS char) `gemmit`," & _
                     "(SELECT COUNT(0) > 0 FROM " & HADBName & ".arzt_has_bs ahb2 LEFT JOIN " & HADBName & ".arzt a2 ON ahb2.arzt_id = a2.idarzt LEFT JOIN " & HADBName & ".arzt_has_genehmigung ahg ON ahg.arzt_id = a2.idarzt LEFT JOIN " & HADBName & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE bs.idbs = ahb2.bs_id AND genehmigung IN ('DMP-DM2_Koordinierender Arzt','DMP-DM2_FA_Diab.bes.qualif.Arzt(SPP)+Koordin. Arzt')) dmpt2," & _
                     "(SELECT COUNT(0) > 0 FROM " & HADBName & ".arzt_has_bs ahb2 LEFT JOIN " & HADBName & ".arzt a2 ON ahb2.arzt_id = a2.idarzt LEFT JOIN " & HADBName & ".arzt_has_genehmigung ahg ON ahg.arzt_id = a2.idarzt LEFT JOIN " & HADBName & ".genehmigung g ON ahg.genehmigung_id = g.idgenehmigung WHERE bs.idbs = ahb2.bs_id AND genehmigung IN ('DMP-DM1_Koordinierender Arzt(SPP)_Erwachsene','DMP-DM1_Koordinierender Arzt(SPP)_Kinder/Jugendliche','DMP-DM1_Koordinierender Arzt_Hausarzt')) dmpt1," & _
                     "`bs`.`aktzeit` AS `aktzeit` " & _
                  "FROM `" & HADBName & "`.`bs` " & _
                  "LEFT JOIN " & HADBName & ".arzt_has_bs ahb ON ahb.bs_id = bs.idbs " & _
                  "LEFT JOIN " & HADBName & ".arzt a ON a.idarzt = ahb.arzt_id " & _
                  "LEFT JOIN " & HADBName & ".`ort` ON `bs`.`Ort_id` = `ort`.`idOrt` " & _
                  "LEFT JOIN " & HADBName & ".`tel` `tel1` ON `tel1`.`bs_id` = `bs`.`idbs` " & _
                  "LEFT JOIN " & HADBName & ".`fax` `fax1` ON `fax1`.`bs_id` = `bs`.`idbs` " & _
                  "LEFT JOIN " & HADBName & ".mail mail ON mail.bs_id = bs.idbs " & _
                  "LEFT JOIN " & HADBName & ".`titel` `t` ON `t`.`idtitel` = `a`.`titel_id` " & _
                  "LEFT JOIN " & HADBName & ".`arzt_has_fachrichtung` `ahf` ON `ahf`.`arzt_id` = `a`.`idarzt` " & _
                  "LEFT JOIN " & HADBName & ".`fachrichtung` `fr` ON `fr`.`idFachrichtung` = `ahf`.`fachrichtung_id` " & _
                  "LEFT JOIN " & HADBName & ".nlart ON a.nlart_id = nlart.idnlart "
'                  "GROUP_CONCAT(a1.nachname) gemmit," & _
'                  "LEFT JOIN " & hadbname & ".arzt_has_bs ahb1 ON ahb1.bs_id = bs.idbs " & _
                  "LEFT JOIN " & hadbname & ".arzt a1 ON a1.idarzt = ahb1.arzt_id AND a1.idarzt <> ahb.arzt_id "

        If irunde = -1 Then
'         rHa.Open sql0 & "WHERE LEFT(bsnr,7) = " & Left$(übwerg(0, runde), 7) & " AND a.nachname = '" & übwerg(1, runde) & "' AND a.vorname = '" & übwerg(2, runde) & "' AND a.obehem=0", _
                  DBCn, adOpenStatic, adLockReadOnly 'haecn
         myFrag rHa, sql0 & "WHERE LEFT(bsnr,7) = " & Left$(übwerg(0, runde), 7) & " AND a.nachname = '" & übwerg(1, runde) & "' AND a.vorname = '" & übwerg(2, runde) & "' AND a.obehem=0", adOpenStatic
        Else
         myFrag rHa, "SELECT h.* FROM `kvaerzte`.`hae" & IIf(irunde = 1, "alt`", "`") & " h WHERE kvnr = '" & Left$(übwerg(0, runde), 2) & "/" & Right$(übwerg(0, runde), 5) & "' AND nachname = '" & übwerg(1, runde) & "' AND vorname = '" & übwerg(2, runde) & "'", adOpenStatic '  -dmpt1 AS j_dmpt1, -dmpt2 AS j_dmpt2, 'haecn
        End If
        If Not rHa.BOF Then If Not IsNull(rHa!kvnu) Then HACngef = True: Exit For
       End If
       If Not HACngef And übwerg(1, runde) <> vNS Then
        Set rHa = Nothing
        If irunde = -1 Then
'         rHa.Open sql0 & "WHERE LEFT(bsnr,7) = " & Left$(übwerg(0, runde), 7) & " AND a.nachname = '" & übwerg(1, runde) & "' AND a.obehem=0", _
                  DBCn, adOpenStatic, adLockReadOnly 'haecn
         myFrag rHa, sql0 & "WHERE LEFT(bsnr,7) = " & Left$(übwerg(0, runde), 7) & " AND a.nachname = '" & übwerg(1, runde) & "' AND a.obehem=0", adOpenStatic
        Else
         myFrag rHa, "SELECT h.* FROM `kvaerzte`.`hae" & IIf(irunde = 1, "alt`", "`") & " h WHERE kvnr = '" & Left$(übwerg(0, runde), 2) & "/" & Right$(übwerg(0, runde), 5) & "' AND nachname = '" & übwerg(1, runde) & "'", adOpenStatic ' -dmpt1 AS j_dmpt1, -dmpt2 AS j_dmpt2, 'haecn
        End If
        If Not rHa.BOF Then If Not IsNull(rHa!kvnu) Then HACngef = True: Exit For
       End If
       If Not HACngef Then
        Set rHa = Nothing
        If irunde = -1 Then
'         rHa.Open sql0 & "WHERE LEFT(bsnr,7) = " & Left$(übwerg(0, runde), 7) & " AND obehem=0", _
                  DBCn, adOpenStatic, adLockReadOnly 'haecn
         myFrag rHa, sql0 & "WHERE LEFT(bsnr,7) = " & Left$(übwerg(0, runde), 7) & " AND obehem=0", adOpenStatic
        Else
         myFrag rHa, "SELECT h.* FROM `kvaerzte`.`hae" & IIf(irunde = 1, "alt`", "`") & " h WHERE kvnr = '" & Left$(übwerg(0, runde), 2) & "/" & Right$(übwerg(0, runde), 5) & "'", adOpenStatic ' -dmpt1 AS j_dmpt1, -dmpt2 AS j_dmpt2,
        End If
        If Not rHa.BOF Then If Not IsNull(rHa!kvnu) Then HACngef = True: Exit For
       End If
       If InfRoh(9, runde) = "" Then InfRoh(9, runde) = übwerg(2, runde) ' vorname
       If InfRoh(14, runde) = "" Then InfRoh(14, runde) = übwerg(1, runde) ' Nachname
       If Trim$(InfRoh(1, runde)) = "Dr.med." Then InfRoh(1, runde) = "Dr.med. " & übwerg(2, runde) & übwerg(1, runde)
      Next irunde
  ' m = 10 + 10 * runde: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
      If HACngef Then
       If Not IsNull(rHa!anrede) Then
        InfRoh(0, runde) = rHa!anrede
       End If
       If Not IsNull(rHa!Titel) Then InfRoh(1, runde) = rHa!Titel Else InfRoh(1, runde) = "Dr.med."
       If Not IsNull(rHa!Vorname) Then InfRoh(1, runde) = InfRoh(1, runde) & " " & rHa!Vorname
       If Not IsNull(rHa!Nachname) Then InfRoh(1, runde) = InfRoh(1, runde) & " " & rHa!Nachname
       If InfRoh(2, runde) = vNS And Not IsNull(rHa.Fields(10)) Then InfRoh(2, runde) = rHa.Fields(10)        ' Straße
       If InfRoh(3, runde) = vNS Then If Not IsNull(rHa!plz) And Not IsNull(rHa!ort) Then InfRoh(3, runde) = rHa!plz & " " & rHa!ort
       If InfRoh(4, runde) = vNS And Not IsNull(rHa!fax1) Then InfRoh(4, runde) = rHa!fax1
       InfRoh(6, runde) = IIf(rHa!dmpt2 <> 0, "X", vNS)
       InfRoh(7, runde) = IIf(rHa!dmpt1 <> 0, "X", vNS)
'       InfRoh(6, runde) = IIF(rHa!j_dmpt2 <> 0, "X", vNS) ' geht auch für null
'       InfRoh(7, runde) = IIF(rHa!j_dmpt1 <> 0, "X", vNS)
'       SET rDMP = Nothing
'       SET rDMP = New ADODB.Recordset
'       myFrag rDMP, "SELECT * FROM `kvaerzte`.`hae` WHERE kvnr = '" & LEFT(übwerg(0, runde), 2) & "/" & Right$(übwerg(0, runde), 5) & "' AND dmpt2 = 1"
'       InfRoh(6, runde) = IIF(rDMP.BOF, vNS, "X")
'       SET rDMP = Nothing
'       SET rDMP = New ADODB.Recordset
'       myFrag rDMP, "SELECT * FROM `kvaerzte`.`hae` WHERE kvnr = '" & LEFT(übwerg(0, runde), 2) & "/" & Right$(übwerg(0, runde), 5) & "' AND dmpt1 = 1"
'       InfRoh(7, runde) = IIF(rDMP.BOF, vNS, "X")
       If LenB(InfRoh(8, runde)) = 0 Then If Not IsNull(rHa!zulg) Then InfRoh(8, runde) = rHa!zulg
       If LenB(InfRoh(9, runde)) = 0 Then If Not IsNull(rHa!Vorname) Then InfRoh(9, runde) = rHa!Vorname
       If LenB(InfRoh(14, runde)) = 0 Then If Not IsNull(rHa!Nachname) Then InfRoh(14, runde) = rHa!Nachname
       InfRoh(10, runde) = übwerg(3, runde)
       If LenB(InfRoh(12, runde)) = 0 Then InfRoh(12, runde) = rHa!kvnu ' Bedingung 25.10.14
       If LenB(InfRoh(13, runde)) = 0 Then If Not IsNull(rHa!tel1) Then InfRoh(13, runde) = rHa!tel1
      End If ' HACngef THEN
    
      gefunden = 0
      If übwerg(1, runde) <> vNS And übwerg(2, runde) <> vNS Then
       Set rhae = Nothing
       myFrag rhae, "SELECT * FROM `hausaerzte` WHERE LEFT(kvnr,7) = '" & Left$(übwerg(0, runde), 7) & "' AND nachname = '" & übwerg(1, runde) & "' AND vorname = '" & übwerg(2, runde) & "'"
       If Not rhae.BOF Then gefunden = True
      End If
      If Not gefunden And übwerg(1, runde) <> vNS Then
       Set rhae = Nothing
       myFrag rhae, "SELECT * FROM `hausaerzte` WHERE LEFT(kvnr,7) = '" & Left$(übwerg(0, runde), 7) & "' AND nachname = '" & übwerg(1, runde) & "'"
       If Not rhae.BOF Then gefunden = True
      End If
      If Not gefunden Then
       Set rhae = Nothing
       myFrag rhae, "SELECT * FROM `hausaerzte` WHERE LEFT(kvnr,7) = '" & Left$(übwerg(0, runde), 7) & "'"
       If Not rhae.BOF Then gefunden = True
      End If
      If gefunden Then
       If InfRoh(0, runde) = vNS Then If Not IsNull(rhae!geschlecht) Then InfRoh(0, runde) = IIf(rhae!geschlecht = "w", "Frau", "Herr")
       If InfRoh(1, runde) = vNS Then
        If Not IsNull(rHa!Titel) Then InfRoh(1, runde) = rHa!Titel Else InfRoh(1, runde) = "Dr.med."
        If Not IsNull(rHa!Vorname) Then InfRoh(1, runde) = InfRoh(1, runde) & " " & rHa!Vorname
        If Not IsNull(rHa!Nachname) Then InfRoh(1, runde) = InfRoh(1, runde) & " " & rHa!Nachname
       End If
       If InfRoh(2, runde) = vNS Then If Not IsNull(rhae!Straße) Then InfRoh(2, runde) = rhae!Straße
       If InfRoh(3, runde) = vNS Then If Not IsNull(rhae!plz) And Not IsNull(rhae!ort) Then InfRoh(3, runde) = rhae!plz & " " & rhae!ort
       If InfRoh(9, runde) = vNS Then If Not IsNull(rhae!Vorname) Then InfRoh(9, runde) = rhae!Vorname
       If InfRoh(14, runde) = vNS Then If Not IsNull(rhae!name) Then InfRoh(14, runde) = rhae!name
       If InfRoh(4, runde) = vNS And Not IsNull(rhae!telefax) Then InfRoh(4, runde) = rhae!telefax
        Select Case rhae!Überschrift
         Case "L": InfRoh(5, runde) = IIf(InfRoh(0, runde) = "Frau", "Liebe", "Lieber") & " " & InfRoh(9, runde) ' nicht: rHae!Geschlecht = "w"
         Case "H": InfRoh(5, runde) = "Hallo " + InfRoh(9, runde)
         Case Else: InfRoh(5, runde) = IIf(InfRoh(0, runde) = "Frau", "Sehr geehrte Frau Kollegin", "Sehr geehrter Herr Kollege")
        End Select
'       InfRoh(6, runde) = IIF(rHae!dmpt2 <> 0, "X", vNS)
'       InfRoh(7, runde) = IIF(rHae!dmpt1 <> 0, "X", vNS)
       If InfRoh(8, runde) = vNS Then If Not IsNull(rhae!zulassungsgebiet) Then InfRoh(8, runde) = rhae!zulassungsgebiet
       InfRoh(10, runde) = übwerg(3, runde)
       If LenB(InfRoh(12, runde)) = 0 Then If Not IsNull(rhae!KVNr) Then InfRoh(12, runde) = rhae!KVNr
       If LenB(InfRoh(13, runde)) = 0 Then If Not IsNull(rhae!telefon) Then InfRoh(13, runde) = rhae!telefon
      End If ' gefunden
      Dim obInsert%
      If Not gefunden Then
       obInsert = True
      Else
       obInsert = False
' übwerg(1, runde) und übwerg(2, runde) sind "", wenn Hausarzt (aus `kvnrue`) aus `namen` im Gegensatz zu Überweiser aus `faelle`
       If Not rhae.BOF Then
        If (rhae!Nachname <> übwerg(1, runde) And übwerg(1, runde) <> vNS) Or (rhae!Vorname <> übwerg(2, runde) And übwerg(1, runde) <> vNS) Then obInsert = True
       End If
      End If
  ' m = 11 + 10 * runde: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
      If obInsert And HACngef Then
' IF obInsert AND NOT rHae.BOF ist falsch.
' wenn rHae.bof dann insert, sonst update
'        Call myEFrag("SET foreign_key_checks = 0")
'        Call myEFrag("DELETE FROM `hausaerzte` WHERE id >= 753")
'        Call myEFrag("SET foreign_key_checks = 1")

'       IF rHae.BOF THEN
        Dim obSchonDa%
        obSchonDa = 0
        Dim NN$, VN$, tit$
        For irunde = 0 To UBound(hains) - 1
         If Not IsNull(rHa!Nachname) Then NN = rHa!Nachname Else NN = ""
         If Not IsNull(rHa!Vorname) Then NN = rHa!Vorname Else NN = ""
         If hains(irunde).kvnu = rHa!kvnu And hains(irunde).Nachname = NN And hains(irunde).Vorname = VN Then
          obSchonDa = True
          Exit For
         End If
        Next irunde
        If Not obSchonDa Then
         Dim rsdop As New ADODB.Recordset
         Set rsdop = Nothing
         If Not IsNull(rHa!Nachname) Then NN = rHa!Nachname Else NN = ""
         If Not IsNull(rHa!Vorname) Then VN = rHa!Vorname Else VN = ""
         myFrag rsdop, "SELECT * FROM `hausaerzte` WHERE kvnr = " & rHa!kvnu & " AND nachname = '" & NN & "' AND vorname = '" & VN & "'"
         If Not rsdop.BOF Then GoTo korrigier
         InsKorr DBCn, "INSERT INTO `hausaerzte`(name, vorname, nachname, anschrift, kvnr, telefon, telefax, e_mail, zulassungsgebiet, arzttyp, " & _
         "`gemeinschaftspraxis mit`" & ", beme, geschlecht,titel,straße,ort,plz,überschrift,dmpt2,dmpt1,zahl,nichtmehr,schwerpunkt,zusatzbezeichnung,bemerkung,sprechstunden) VALUES('" & IIf(LenB(übwerg(2, runde)) = 0 And LenB(übwerg(1, runde)) <> 0, übwerg(1, runde), rHa!anrede & " " & IIf(rHa!Titel <> vNS And Not IsNull(rHa!Titel), rHa!Titel, "Dr.med.") & " " & IIf(LenB(übwerg(1, runde)) = 0, rHa!Vorname, übwerg(2, runde)) & " " & IIf(LenB(übwerg(1, runde)) = 0, rHa!Nachname, übwerg(1, runde))) & _
         "','" & IIf(LenB(übwerg(1, runde)) = 0, rHa!Vorname, übwerg(2, runde)) & "','" & IIf(LenB(übwerg(1, runde)) = 0, rHa!Nachname, übwerg(1, runde)) & "','" & rHa!Straße & ", " & rHa!plz & " " & rHa!ort & "','" & rHa!kvnu & "','" & rHa!tel1 & "','" & rHa!fax1 & "','" & rHa!email & "','" & rHa!zulg & "','" & rHa!arzttyp & "','" & rHa!gemmit & "','" & rHa!beme & "','" & IIf(rHa!anrede = "w", "Frau", "Herr") & "','" & rHa!Titel & "','" & rHa!Straße & "','" & rHa!ort & "','" & rHa!plz & "',''," & IIf(InfRoh(6, runde) = "X", "1", "0") & "," & IIf(InfRoh(7, runde) = "X", "1", "0") & ",0,0,'','','','')", rAf
         hains(UBound(hains)).kvnu = rHa!kvnu
         If IsNull(rHa!Nachname) Then hains(UBound(hains)).Nachname = vNS Else hains(UBound(hains)).Nachname = rHa!Nachname
         If IsNull(rHa!Vorname) Then hains(UBound(hains)).Vorname = vNS Else hains(UBound(hains)).Vorname = rHa!Vorname
         ReDim Preserve hains(UBound(hains) + 1)
        End If
      ElseIf Not rhae.BOF And Not rHa.BOF Then
        If Not IsNull(rHa!Nachname) Then NN = rHa!Nachname Else NN = ""
        If Not IsNull(rHa!Vorname) Then VN = rHa!Vorname Else VN = ""
        If Not IsNull(rHa!Titel) Then tit = rHa!Titel Else tit = "Dr.med."
        If rhae!name <> rHa!anrede & " " & tit & " " & VN & " " & VN Or _
           rhae!Vorname <> VN Or rhae!Nachname <> NN Or _
           rhae!anschrift <> rHa.Fields(10) & ", " & rHa!plz & " " & rHa!ort Or _
           rhae!telefon <> rHa!tel1 Or rhae!telefax <> rHa!fax1 Or rhae!e_mail <> rHa!email Or _
           rhae!zulassungsgebiet <> rHa!zulg Or rhae!arzttyp <> rHa!arzttyp Or rhae.Fields("gemeinschaftspraxis mit") <> rHa!gemmit Or rhae!beme <> rHa!beme Or _
           rhae!dmpt2 <> rHa!dmpt2 Or rhae!dmpt1 <> rHa!dmpt1 _
           Then
korrigier:
        myEFrag "UPDATE `hausaerzte` SET name = '" & rHa!anrede & " " & IIf(rHa!Titel <> vNS And Not IsNull(rHa!Titel), rHa!Titel, "Dr.med.") & " " & rHa!Vorname & " " & rHa!Nachname & "', vorname = '" & rHa!Vorname & "', nachname = '" & rHa!Nachname & "', anschrift = '" & rHa.Fields(10) & ", " & rHa!plz & " " & rHa!ort & "', telefon = '" & rHa!tel1 & "', telefax = '" & rHa!fax1 & "', e_mail = '" & rHa!email & "', zulassungsgebiet = '" & rHa!zulg & "', arzttyp = '" & rHa!arzttyp & "', `gemeinschaftspraxis mit` = '" & rHa!gemmit & "', beme = '" & rHa!beme & "', geschlecht = '" & IIf(rHa!anrede = "w", "Frau", "Herr") & "', titel = '" & rHa!Titel & "', straße = '" & rHa.Fields(10) & "',ort = '" & rHa!ort & "', plz = '" & rHa!plz & "', überschrift = " & "''" & ",dmpt2 = '" & IIf(InfRoh(6, runde) = "X", "1", "0") & "', dmpt1 = '" & IIf(InfRoh(7, runde) = "X", "1", "0") & "', zahl = " & "0" & ",nichtmehr = " & "0" & ",schwerpunkt = " & "''" & ",zusatzbezeichnung = " & "''" & ",bemerkung = " & "''" & _
                     ",sprechstunden  = " & "''" & " WHERE kvnr = '" & rHa!kvnu & "' AND nachname = '" & rHa!Nachname & "' AND vorname = '" & rHa!Vorname & "'", rAf
       End If
'       GoTo hier:
      End If
     End If ' übwerg(0,runde)<> vns
     If InfRoh(5, runde) = vNS Then InfRoh(5, runde) = IIf(InfRoh(0, runde) = "Frau", "Sehr geehrte Frau Kollegin", IIf(InfRoh(0, runde) = "Herr", "Sehr geehrter Herr Kollege", "Sehr geehrte Frau Kollegin, sehr geehrter Herr Kollege"))
     InfRoh(11, runde) = IIf(InStrB(LCase$(InfRoh(8, runde)), "intern") <> 0 Or InStrB(LCase$(InfRoh(8, runde)), "innere") <> 0 Or InStrB(LCase$(InfRoh(8, runde)), "allg") <> 0 Or InStrB(LCase$(InfRoh(8, runde)), "prakt") <> 0 Or InfRoh(8, runde) = "Arzt" Or InfRoh(8, runde) = "Ärztin" Or InfRoh(8, runde) = "Hausarzt" Or InfRoh(8, runde) = "Hausärztin", "X", vNS)
  ' m = 12 + 10 * runde: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
  ' m = 13 + 10 * runde: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
  ' m = 14 + 10 * runde: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
    Next runde ' runde = 0 To UBound(übwerg, 2)
    ' m = 100: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
    Dim pos&, i%, aktrunde%, gelöscht%
    pos = UBound(InfRoh, 2)
' zuletzt kommen die Überweiser, die nicht als Allgemeinarzt/Internist sind und nicht als Hausarzt eingetragen
    For runde = UBound(InfRoh, 2) To 0 Step -1
     If InfRoh(11, runde) = vNS And InStrB(InfRoh(10, runde), "HA") = 0 Then
      For i = 0 To UBound(InfRoh, 1)
       infos(i, pos) = InfRoh(i, runde)
'       InfRoh(i, runde) = vns
      Next i
      pos = pos - 1
     End If
    Next runde
    ' m = 101: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
    aktrunde = UBound(InfRoh, 2) + 1 ' der letzte hier berücksichtigte
' dann kommen bis zur 3. Stelle die Überweiser, die als Allgemeinarzt/Internist sind aber nicht als Hausarzt eingetragen
    For runde = UBound(InfRoh, 2) To 0 Step -1
     If pos <= IIf(obHAPrio, 0, 1) Then Exit For
     If InfRoh(11, runde) = "X" And InStrB(InfRoh(10, runde), "HA") = 0 Then
      gelöscht = 0
      For irunde = UBound(InfRoh, 2) To 0 Step -1
       If irunde <> runde Then
        If InfRoh(1, irunde) = InfRoh(1, runde) Then ' doppelter Arzt
         For i = 0 To UBound(InfRoh, 1)
          InfRoh(i, runde) = vNS
         Next i
         gelöscht = True
         Exit For
        End If
       End If
      Next irunde
      If Not gelöscht Then ' war ja oben so, also falsch nicht gelöscht
       For i = 0 To UBound(InfRoh, 1)
        infos(i, pos) = InfRoh(i, runde)
        InfRoh(i, runde) = vNS ' schon eingetragen
       Next i
      End If
      pos = pos - 1
     End If
     aktrunde = runde
    Next runde
    ' m = 102: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
' dann kommen Hausärzte, falls bisher nicht berücksichtigt
    For runde = UBound(InfRoh, 2) To aktrunde Step -1
     If InStrB(InfRoh(10, runde), "HA") <> 0 Then
      gelöscht = 0
      For irunde = UBound(InfRoh, 2) To 0 Step -1
       If irunde <> runde Then
        If InfRoh(1, irunde) = InfRoh(1, runde) And Not (InStrB(InfRoh(10, irunde), "HA") = 0 And InStrB(InfRoh(10, runde), "HA") <> 0) Then ' doppelter Arzt, Einschränkung 8.2.10
         For i = 0 To UBound(InfRoh, 1)
          If i = 10 Then
           InfRoh(i, irunde) = InfRoh(i, irunde) & "," & InfRoh(i, runde)
          End If
          InfRoh(i, runde) = vNS
         Next i
         gelöscht = True
         Exit For
        End If
       End If
      Next irunde
      If Not gelöscht Then ' war ja oben so, also falsch nicht gelöscht
       For i = 0 To UBound(InfRoh, 1)
        infos(i, pos) = InfRoh(i, runde)
        InfRoh(i, runde) = vNS
       Next i
      End If
      pos = pos - 1
'      Exit For ' auskommentiert 1.11.14
     End If ' InStrB(InfRoh(10, runde), "HA") <> 0 Then
    Next runde
    ' m = 103: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
' dann der Rest in der Reihenfolge
    For runde = aktrunde - 1 To 0 Step -1
     If InfRoh(11, runde) = "X" Or InStrB(InfRoh(10, runde), "HA") <> 0 Then
      gelöscht = 0
      For irunde = UBound(InfRoh, 2) To 0 Step -1
       If irunde <> runde Then
        If InfRoh(1, irunde) = InfRoh(1, runde) And Not (InStrB(InfRoh(10, irunde), "HA") = 0 And InStrB(InfRoh(10, runde), "HA") <> 0) Then ' doppelter Arzt, Einschränkung 28.8.08
         For i = 0 To UBound(InfRoh, 1)
          If i = 10 Then
           InfRoh(i, irunde) = InfRoh(i, irunde) & "," & InfRoh(i, runde)
          End If
          InfRoh(i, runde) = vNS
         Next i
         gelöscht = True
         Exit For
        End If
       End If
      Next irunde
      If Not gelöscht Then  ' war ja oben so, also falsch nicht gelöscht
       For i = 0 To UBound(InfRoh, 1)
        infos(i, pos) = InfRoh(i, runde)
        InfRoh(i, runde) = vNS
       Next i
      End If
      pos = pos - 1
     End If ' InfRoh(11, runde) = "X" Or InStrB(InfRoh(10, runde), "HA") <> 0 Then
    Next runde
    ' m = 104: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
' Wenn erste Position leer, schaut's unschön aus
    If infos(1, 0) = vNS Then
     If UBound(infos, 2) > 0 Then
      If infos(1, 1) <> vNS Then
       For i = 0 To UBound(infos, 1)
        infos(i, 0) = infos(i, 1)
        infos(i, 1) = vNS
       Next i
      Else
       If UBound(infos, 2) > 1 Then
        If infos(1, 2) <> vNS Then
         For i = 0 To UBound(infos, 1)
          infos(i, 0) = infos(i, 1)
          infos(i, 1) = vNS
         Next i
        End If
       End If
      End If
     End If
    End If
    ' m = 105: Tj(m) = Timer ': for p = 0 To m - 1: Tj(m) = Tj(m) - Tj(p): Next p
    ' 28.8.21
    For i = UBound(infos, 2) - 1 To 0 Step -1
     If infos(9, i) = "" And infos(14, i) = "" And infos(9, i + 1) <> "" And infos(14, i + 1) <> "" Then ' wenn Vor- und Nachnahme fehlen
      Dim ii&
      For ii = i To UBound(infos, 2) - 1
       Dim j&
       For j = 0 To UBound(infos, 1)
        infos(j, ii) = infos(j, ii + 1)
       Next j
      Next ii
      ReDim Preserve infos(UBound(infos, 1), UBound(infos, 2) - 1)
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in gethausarzt1/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' gethausarzt1

' aufgerufen in VorhandeneBriefe_Click
Function doVorhandene&()
 Dim Fil As File, pid&, pos&, p2&, BriefZiel$
 Dim VorMüll$, FilName$, FilPath$
 On Error GoTo fehler
 BriefZiel = InputBox("Verzeichnis:", "wo sollen vorhandene Briefe korrigiert werden?", AutoBriefZiel)
 If BriefZiel <> vNS Then
  For Each Fil In FSO.GetFolder(BriefZiel).Files
   FilPath = Fil.path
   pos = InStr(FilPath, "PID")
'   Debug.Print FilPath
   If pos > 0 Then
    p2 = InStr(pos, FilPath, ",")
    If p2 > 0 Then
     pid = CLng(Mid$(FilPath, pos + 4, p2 - pos - 4))
     If pid > 0 Then
      VorMüll = FilPath & "zulöschen"
      FilName = Fil.name
      Name Fil.path As VorMüll
      Forms(0).Ausgeb "Erneuere: " & FilName, True
      Call tubriefStandalone(pid, True, BriefZiel)
      If BrichAb Then Exit For
      Kill VorMüll
      doVorhandene = doVorhandene + 1
nächstedatei:
     End If
    End If
   End If
  Next Fil
 End If
 Exit Function
fehler:
  Dim FNr&, FLastDLLError&, FSource$, FDescr$
  FNr = Err.Number
  FLastDLLError = Err.LastDllError
  FSource = Err.source
  FDescr = Err.Description
  Call meld("Achtung Fehler beim Erneuern von: " & FilPath, 0)
  PiepKurz
  If FNr = 75 Then Resume nächstedatei
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(FNr) + vbCrLf + "LastDLLError: " + CStr(FLastDLLError) + vbCrLf + "Source: " + IIf(IsNull(FSource), vNS, CStr(FSource)) + vbCrLf + "Description: " + FDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in doVorhandene/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doVorhandene

' in DuplexKontrollieren_Click
Function doDuplexkontrollieren()
 Dim rs As New ADODB.Recordset, begD As Date, sql$, begDStr$
 begDStr = InputBox("Ab welchem Datum?")
 If IsDate(begDStr) Then
  begD = CDate(begDStr)
  sql = "SELECT pat_id, zeitpunkt, MIN(name) DokName, quelldatum, COUNT(0) zahl FROM briefe b WHERE name LIKE '%SonoBild%' AND quelldatum >= " & DatFor_k(begD) & " GROUP BY pat_id, zeitpunkt,DATE(quelldatum) ORDER BY quelldatum"
'  rs.Open sql, DBCn, adOpenDynamic, adLockReadOnly
  myFrag rs, sql
  TabAusgeb rs, Lese, True, , , , , , "Duplexkontrolle " & begDStr & " - " & Date
'  Do While Not rs.EOF
'   Forms(0).Ausgeb Right$(Space$(5) & rs!Pat_id, 5) & " " & rs!Zeitpunkt & " " & Right$(Space$(3) & rs!Zahl, 3) & " " & rs!DokName & " " & rs!Quelldatum, True
'   rs.Move 1
'  Loop
 End If
End Function ' doDuplexkontrollieren

Function doBriefeBerichtspflicht()
 Dim QDat$, dszahl&, angefangen%, erg&, ergS$, fdt As Date
 On Error Resume Next
' QDat = getLDatei(uverz, "Listenausgabe_EBM-Ziffern*.xls")
' IF Err.Number <> 0 THEN
'  MsgBox "Fehler beim Einladen der Datei " & uVerz & "Listenausgabe_EBM-Ziffern*.xls"
'  Exit Function
' END IF
 QDat = uVerz & "Listenausgabe_Listen.xls"
 ergS = Dir(QDat)
 If ergS <> vNS Then
  fdt = FileDateTime(QDat)
  erg = MsgBox("Die Datei " & QDat & " stammt von: " & fdt & "." & vbCrLf & "Wollen Sie zu allen dort enthaltenen Patienten Briefe schreiben und abrechnen?", vbOKCancel, "Rückfrage")
  If erg = vbNo Then Exit Function
 Else
  MsgBox ("Datei " & QDat & " nicht gefunden, breche ab!")
  Exit Function
 End If
 On Error GoTo fehler
' aus FUNCTION EmailsImport(EmDatei$)
 Dim con As New ADODB.Connection  ' Connection
 Dim rNa As New ADODB.Recordset
 Dim rEx As New ADODB.Recordset
 Dim rX As New ADOX.Catalog
 Dim BDT As New BDTSchreib
' Const EmDatei$ = pverz & "Patientenübergreifendes\Emails.xls" ' Excel-Datei mit Suche aus Turbomed "*@*"
 On Error GoTo fehler
 If LenB(DBCn) = 0 Or DBCn = vNS Then
'   Call frm.ConstrFestleg(2)
   Call acon(quelleT, qDtb)
 End If
 If QDat <> vNS Then
 con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=""Excel 8.0;HDR=No;IMEX=1"";Data Source=" & QDat & ";" ' TABLE=Adressen$"
 Dim runde%, i%, zFeld$, lFeld$, obAnfang%, pNr%, pRoh
  rX.ActiveConnection = con
  rEx.Open "`" & rX.Tables(rX.Tables.COUNT - 1).name & "`", con ' Hier Excel, nicht lies.obmysql = 0!
  Do While Not rEx.EOF
'  Debug.Print runde
   If obAnfang Then
'    ON Error Resume Next
     pRoh = rEx.Fields(0)
     If InStrB(rEx.Fields(2), "Es ist kein Bericht vorhanden") > 0 Then
' Hier jetzt Vorname, Name und Geburtsdatum extrahieren, um den Patienten zu finden!
       Dim VN$, NN$, gd$, gesn$, leid$
       Dim p1%, p2%, p3%
       p1 = InStr(pRoh, " ")
       p2 = InStr(pRoh, "(")
       p3 = InStr(pRoh, " - ")
       VN = Left$(pRoh, p1 - 1)
       gesn = REPLACE$(Left$(pRoh, p2 - 2), "zzz", vNS)
       NN = Mid$(pRoh, p1 + 1, InStr(p1 + 1, pRoh, " ") - p1 - 1)
       gd = Mid$(pRoh, p2 + 1, InStr(pRoh, ")") - p2 - 1)
       leid = Mid$(pRoh, p3 + 3, InStr(pRoh, ";") - 3 - p3)
       dszahl = dszahl + 1
       If NN Like "zzz*" Then NN = Mid$(NN, 4)
       Set rNa = Nothing
       sql = "SELECT * FROM `namen` WHERE CONCAT(titel¡iIF(titel='','',' ')¡nvorsatz¡iIF(nvorsatz='','',' ')¡vorname¡' '¡nachname) = '" & gesn & "' AND gebdat = " & DatFor_k(gd) & ""
       If Forms(0).obMySQL Then
        sql = REPLACE$(REPLACE$(sql, "¡", ","), "iIF(", "IF(")
       Else
        sql = REPLACE$(REPLACE$(sql, "concat", vNS), "¡", " & ")
       End If
       myFrag rNa, "SELECT COUNT(0) FROM (" & sql & ") i", adOpenStatic, DBCn, adLockOptimistic
       If rNa.Fields(0) <> 1 Then
        MsgBox "Falsche Zahl an Datensätzen für " & VN & " " & NN & ", geb. " & gd & ": " & rNa.Fields(0)
       Else
'        Set rNa = Nothing
        myFrag rNa, sql, adOpenStatic, DBCn, adLockOptimistic
        If Not angefangen Then
         If Not BDT.Start(hVerz, "Leist", 0) Then
          Exit Function
         End If
         Call BDT.BDTKopf
         Call BDT.ImportFolderHerricht
'          Call LeistungsExport0
          angefangen = True
        End If
'        Call LeistungsExport1(BDT, rNa!Pat_id, "01601", CDate(leid), CDate("18:00"))
        Call LeistungsExport1(BDT, rNa!Pat_ID, "40110", CDate(leid), CDate("18:00"), , 0)
        Call tubriefStandalone(rNa!Pat_ID, True)
       End If
     End If
   ElseIf Not IsNull(rEx.Fields(0)) And Not IsNull(rEx.Fields(1)) And Not IsNull(rEx.Fields(2)) Then
     obAnfang = True
   End If
  runde = runde + 1
  rEx.MoveNext
  Loop
 End If
 On Error Resume Next
 rEx.Close
 Forms(0).Ausgeb "Fertig mit Nachtragen aus '" & QDat & "' von " & dszahl & " Datensätzen ", True
 If angefangen Then
'  Close #310
  MsgBox "Datei " & BDT.DMPImp & " neu mit den Leistungen zu den Briefen in " & QDat & " erstellt!"
 End If
 Exit Function
fehler:
  Dim FNr&, FLastDLLError&, FSource$, FDescr$
  FNr = Err.Number
  FLastDLLError = Err.LastDllError
  FSource = Err.source
  FDescr = Err.Description
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(FNr) + vbCrLf + "LastDLLError: " + CStr(FLastDLLError) + vbCrLf + "Source: " + IIf(IsNull(FSource), vNS, CStr(FSource)) + vbCrLf + "Description: " + FDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in doBriefeBerichtspflicht/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doBriefeBerichtspflicht

'lfaelle: SELECT *
'FROM `faelle` INNER JOIN `SELECT MIN(pat_id) AS pid, MAX(bhfb) AS bb FROM `faelle` GROUP BY pat_id`. AS sel ON (`faelle`.`bhfb`=sel.bb) AND (`faelle`.`pat_id`=sel.pid)
'ORDER BY `pat_id`;
' Vorlage: getDokPfad("Vorlagen") + "\AccessBrief.dot" ' xVerz & \vorlagen\AccessBrief.dot (s.u.)
#If mitab Then
Public Sub tu_brief(frm As Form)
 ' frm.anaRS!Pat_id
 Call tubriefStandalone(Pat_ID, 0)
End Sub ' tu_brief
#End If

'Function stil1(ByRef stil As CString, Optional SZ% = 24, Optional van%)
'' Set stil = New CString
' stil.Append "<w:rPr><w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial"" w:cs=""Arial""/>"
' If van Then stil.Append "<w:vanish/>"
' stil.AppVar Array("<w:sz w:val=""", SZ, """/>")
' stil.Append "<w:lang w:val=""en-GB""/></w:rPr>"
'End Function ' stil

Function stil$(Optional mitcs As Byte, Optional SZ% = 24, Optional mitszcs As Byte, Optional bold As Byte, _
               Optional ital As Byte, Optional van As Byte, Optional lang$, Optional pos%, Optional underl As Byte)
' Set stil = New CString
 stil = "<w:rPr><w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial""" & IIf(mitcs, " w:cs=""Arial""", "") & "/>"
 If bold And 1 Then stil = stil & "<w:b/>"
 If bold And 2 Then stil = stil & "<w:bCs/>"
 If ital <> 0 Then stil = stil & "<w:i/><w:iCs/>"
 If van <> 0 Then stil = stil & "<w:vanish/>"
 If pos <> 0 Then stil = stil & "<w:position w:val=""" & pos & """/>"
 If SZ <> 0 Then
  stil = stil & "<w:sz w:val=""" & SZ & """/>"
  If mitszcs Then stil = stil & "<w:szCs w:val=""" & SZ & """/>"
 End If
 If underl <> 0 Then stil = stil & "<w:u w:val=""single""/>"
 If lang$ <> "" Then stil = stil & "<w:lang w:val=""" & lang & """/>"
 stil = stil & "</w:rPr>"
End Function ' stil

Function stiltxt$(Text$, Optional bold As Byte, Optional ital As Byte, Optional underl As Byte)
 stiltxt = "<w:r>" & stil(, , , bold, ital, , , , underl) & "<w:t xml:space=""preserve"">" & Text$ & "</w:t></w:r>"
' stiltxt = "<w:r>" & stil(, , , bold, ital, , , , underl) & "<w:t>" & Text$ & "</w:t></w:r>"
End Function ' stiltxt

Function htxt$(Text$)
 htxt = "<w:t>" & Text & "</w:t>"
End Function ' htxt

'' bei 100000 Anhängungen geht "" & "" langsamer as CString, bei weniger deutlich umgekehrt
'Public Function vstile()
' Dim i&, ST As New CString, stl$
' Debug.Print Now()
' For i = 0 To 100000
''  Call stil1(st)
'  stl = stl & stil
' Next i
' Debug.Print Now()
'End Function

Function bookmS(bname$)
 bookmS = "<w:bookmarkStart w:id=""" & bmnr & """ w:name=""" & bname$ & """/>"
End Function ' dobookm(id&, bname$)

Function bookmE()
 ag.AppVar Array("<w:bookmarkEnd w:id=""", bmnr, """/>")
 bmnr = bmnr + 1
End Function ' bookmE

' aufgerufen in tubriefStandalone
Sub bookm(bmname$, Inh$, Optional obvanish As Byte, Optional obbold As Byte)
 ag.AppVar Array(bookmS(bmname), "<w:r>", stil(, , 1, obbold, , obvanish, "en-GB"))
' If Inh$ <> "" Then ag.AppVar Array("<w:t", IIf(Inh = " ", " xml:space=""preserve""", ""), ">", Inh, "</w:t>")
  If Inh$ <> "" Then ag.AppVar Array("<w:t>", Inh, "</w:t>")
  ag.AppVar Array("</w:r>", bookmE)
End Sub ' bookm

' aufgerufen in tubriefStandalone
Sub agpreserve()
' ag.Append "<w:r><w:rPr><w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial""/><w:sz w:val=""24""/><w:szCs w:val=""24""/><w:lang w:val=""en-GB""/></w:rPr><w:t xml:space=""preserve""> </w:t></w:r>"
 ag.Append "<w:r><w:rPr><w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial""/><w:sz w:val=""24""/><w:szCs w:val=""24""/><w:lang w:val=""en-GB""/></w:rPr><w:t> </w:t></w:r>"
End Sub ' agpreserve

' aufgerufen in tubriefStandalone
Sub agtabu()
 ag.Append "<w:r><w:rPr><w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial""/><w:sz w:val=""24""/><w:lang w:val=""en-GB""/></w:rPr><w:tab/></w:r>"
End Sub ' agtabu

Function doagtab$(Optional pos$ = "5387")
    doagtab = "<w:pPr><w:tabs><w:tab w:val=""left"" w:pos=""" & pos & """/></w:tabs>"
End Function ' doagtab

Sub agtab2()
 ag.Append doagtab & stil(, , , , , , "en-GB") & "</w:pPr>"
End Sub ' agtab2

Function agabsa$()
 Static absnr&
 agabsa = "<w:p w14:paraId=""" & Right$("0000000" & Hex(absnr), 8) & """ w14:textId=""77777777"" w:rsidR=""00000000"" w:rsidRDefault=""" & donr & """>"
 absnr = absnr + 1
End Function ' agabsa

' in BriefImport_Click
Public Sub tubriefStandalone(pid&, obStumm%, Optional Zielverz$, Optional Vorlage$, Optional nurLabor% = False, Optional briefneu% = False, Optional nichtherricht%) ' Brief schreiben
 Dim Pat_ID$, myRange, Docu, Inh, dc As Object ' Word.Document, wegen unbekannter Wordversion als Object
 Dim raHa As New ADODB.Recordset
 Dim raAn As New ADODB.Recordset
 Dim raHae As New ADODB.Recordset
 Dim rsNa As New ADODB.Recordset
 Dim obHAimDMP%
 Dim j&, pos%, VBuch$
' dim fs
 Dim runde%, KVNr$
 Dim VorDat As Date
 Dim ÜWNr$, ÜWeingef As Boolean
    ÜWNr = vNS
    ÜWeingef = False
 Dim sverz$
  
 If Date = #11/20/2011# Then nurLabor = True
 On Error GoTo fehler
 Pat_ID = CStr(pid)
'  sverz = pverz
 If Zielverz = vNS Then
  sverz = BriefZiel
  If obStumm Then
'    sverz$ = sverz & "unkorrigiert\"
   sverz = AutoBriefZiel
  End If
 Else ' Zielverz = vNS Then
  sverz = Zielverz
 End If ' Zielverz = vNS Then else
'  lies.obmysql = False
'  Call Zinit(Lies.obMySQL)
 nzw = vbCr
  
'  SET rsNa = TabÖff("Namen", "pat_ID")
'  rsNa.Seek "=", raAn!Pat_id
 Dim NachNa$, VorNa$, GName$, G1Name$, tit$, GebDat As Date, dieder$, lbeh As Date
 myFrag rsNa, "SELECT n.*,COALESCE(titel,'') tit, CONCAT(IF(geschlecht='m','Herrn','Frau'),' ',gesname(pat_id),', *',DATE_FORMAT(gebdat,'%e.%c.%y')) gname,IF(geschlecht='m','der','die') dieder, lbeh FROM namenlb n WHERE pat_id = " & Pat_ID
 If rsNa.State = 0 Then
  MsgBox "Namen-Tabelle nicht abfragbar. Es könnten z.B. die Indices von einlesen korrupt sein."
  Exit Sub
 End If ' rsNa.State = 0 Then
 If rsNa.EOF Then Exit Sub
 NachNa = rsNa!Nachname
 VorNa = rsNa!Vorname
 GName = rsNa!GName
 GebDat = rsNa!GebDat
 tit = rsNa!tit
 dieder = rsNa!dieder
 lbeh = rsNa!lbeh
' G1Name = rsNa!G1Name
 myFrag raAn, "SELECT COALESCE(GesName(Pat_id),'') GesName, COALESCE(`diabetes seit`,0) dmseit FROM `anamnesebogen` a WHERE pat_id = " & Pat_ID, adOpenStatic
 Dim gesname$, dmseit$
 gesname = raAn!gesname
 dmseit = raAn!dmseit
 syscmd acSysCmdSetStatus, "Erstelle Brief für " & gesname & ": 1) Hausarzt ..."
 Dim rFa() As Faelle
' Dim rKv1() As kvnrue
 Dim FaxNr$, infos$() ' Frau/Herrn, Vorn+Nachn, Straße, PLZ+Ort, Faxnr, S.g./Liebe, DMPTyp2, DMPTyp1
 Call getHausarzt1(infos(), rFa, rKv, 1, Pat_ID, 0, 2) ', True)
  
'  SET rHa = TabÖff("Hausaerzte", "KVNR")
 If LenB(Vorlage) = 0 Then Vorlage = "AccessBrief.dot"
 If Right$(LCase$(Vorlage), 4) = ".dot" Then If Left$(LCase$(Vorlage), 11) = "accessbrief" Then VBuch = Mid$(Vorlage, 12, Len(Vorlage) - 15)
 Dim DiagTab() As CString, DString$
 DString = DiagString(Pat_ID, DiagTab, VorDat, obBrief:=True, dmseit:=dmseit)
 If briefneu Then
  Dim aname$
  aname = ArBName$(sverz, NachNa, VorNa, Pat_ID$, infos$())
  If Not nichtherricht Then
   Dim oSh As New IWshShell_Class
'  SET oSh = New IWshShell_Class
   oSh.rUn "cmd /c ""xcopy " & vVerz & "exp8\ " & vVerz & "exp9\ /s /y /h /r /k /c && move " & vVerz & "exp9\word\media\image" & VBuch & ".jpeg " & vVerz & "\exp9\word\media\1.jpeg && del " & vVerz & "exp9\word\media\image*.jpeg && move " & vVerz & "exp9\word\media\1.jpeg " & vVerz & "exp9\word\media\image1.jpeg""", 0, True
   Open vVerz & "exp9\docProps\app.xml" For Output As #51
   Print #51, app1 & Vorlage & app2
   Close #51
   Open vVerz & "exp9\docProps\core.xml" For Output As #51
   Print #51, core1 & Environ("username") & core2 & Environ("username") & core3 & Format(Now(), "YYYY-mm-ddThh:MM:ssZ") & core4 & Format(Now(), "YYYY-mm-ddThh:MM:ssZ") & core5
   Close #51
   Open vVerz & "exp9\word\_rels\settings.xml.rels" For Output As #51
   Print #51, settings1 & "c:\turbomed\vorlagen\" & Vorlage & settings2
   Close #51
  ' in .run wird das " durch "" escaped, in powershell durch \"
'  oSh.run "powershell ""$dt=\""v:\exp9\word\endnotes.xml\"",\""v:\exp9\word\footnotes.xml\"";$dth=\""v:\exp9\word\header1.xml\"";$nrd=\""\\" & LiName & "\daten\eigene dateien\programmierung\dateilesen\dzahl.txt\"";$nr=[int](get-content -path $nrd)+1;set-content -path $nrd $nr;$repl='${1}'+'{0:X8}' -f $nr+'$2';$qla='(.*w:rsidRDefault=\"")[0-9A-F]*(\""><w:';$qls=$qla+'r><w:';foreach ($dta in $dt){((get-content -path $dta) -replace \""${qls}s.*)\"",$repl) -replace \""${qls}c.*)\"",$repl|set-content -path $dta;};(get-content -path $dth)|foreach-object {$_ -replace \""${qla}p.*)\"",$repl}|set-content -path $dth;""", 0, True
   Dim dzahl$
   dzahl = uVerz & "Programmierung\Dateilesen\dzahl.txt"
   oSh.rUn "powershell ""$vz=\""" & vVerz & "exp9\word\\\"";$dt=@($vz+\""endnotes.xml\"";$vz+\""footnotes.xml\"";$vz+\""document.xml\"";$vz+\""settings.xml\"");$anr=[string](get-content -path $dt[0]) -replace '.*w:rsidRDefault=\""([0-9A-F]*)\"".*','$1';$nrd=\""" & dzahl & "\"";$nr=[int](get-content -path $nrd)+1;set-content -path $nrd $nr;$nrs='{0:X8}' -f $nr;foreach ($dta in $dt){(get-content -path $dta) -replace $anr, $nrs|set-content -path $dta;};""", 0, True
   Dim docid$
   Randomize
   docid = Right$("0000000" & Hex(Rnd * (16 ^ 7 - 1)), 7)
   Randomize
   docid = docid & Format$(Right$("0000000" & Hex(Rnd * (16 ^ 7 - 1)), 7), "&-&&&&-&&")
   Randomize
   docid = docid & Format$(Right$("0000000" & Hex(Rnd * (16 ^ 7 - 1)), 7), "&&-&&&&-&")
   Randomize
   docid = docid & Right$("0000000" & Hex(Rnd * (16 ^ 7 - 1)), 7)
   Randomize
   docid = docid & Right$("0000000" & Hex(Rnd * (16 ^ 7 - 1)), 4)
   oSh.rUn "powershell ""$dt=\""" & vVerz & "exp9\word\settings.xml\"";$repl='${1}'+'" & docid & "'+'$2';$qla='(.*{)[0-9A-F-]*(}.*)';(get-content -path $dt) -replace $qla,$repl|set-content -path $dt;""", 0, True
  
   donr = Right$("00000000" & Hex(ReadFile(dzahl)), 8)
   Open vVerz & "exp9\word\header1.xml" For Output As #51
   Print #51, header1 & donr & header2 & zuh(GName) & header3 & Format(Now(), "dd.mm.yy") & header4 & donr & header5
   Close #51
   Open vVerz & "exp9\word\document.xml" For Output As #51
'  If UBound(infos, 2) = 0 Then ReDim Preserve infos(UBound(infos, 1), 1)
'  Print #51, zuh(doc1 & kopf0 & doc2 & donr & doc3 & donr & doc4 & donr & doc5 & donr & doc6 & donr & doc7 & donr & doc8 & donr & doc9 & Format(Now(), "dd.mm.yyyy") & doc10 & donr & doc11 & donr & doc11 & nr & doc12 & donr & doc13 & donr & doc14 & _
  IIf(infos(0, 0) = "Herr", "Herrn", infos(0, 0)) & doc15 & IIf(infos(0, 1) = "Herr", "Herrn", infos(0, 1)) & doc16 & nr & doc17 & infos(1, 0) & doc18 & infos(1, 1) & doc19 & nr & doc20 & infos(2, 0) & doc21 & infos(2, 1) & doc22 & nr & doc23 & nr & doc24 & infos(3, 0) & doc25 & infos(3, 1) & doc26 & nr & doc27 & nr & doc28 & nr & doc28 & nr & doc28 & nr & doc29 & _
  "Liebe Kolleginnen </w:t></w:r></w:p>" & _
  doce)
'  Dim doc1a$, doc1b$
  
   ag.AppVar Array(doc1, kopf0, "<w:body>", agabsa, vors, doc2, agabsa, vors, doc3, agabsa, vors, doc4, agabsa, vors, zuh(doc5), agabsa, doc6, agabsa, doc7, agabsa, doc8, Format(Now(), "dd.mm.yyyy"), doc9, agabsa, doc10, agabsa, zuh(doc11), agabsa, doc12, agabsa)
   ag.Append "<w:pPr><w:tabs><w:tab w:val=""left"" w:pos=""5387""/></w:tabs><w:spacing w:before=""140""/><w:rPr><w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial""/><w:sz w:val=""24""/><w:lang w:val=""en-GB""/></w:rPr></w:pPr>"
   bookm "ha1a", IIf(infos(0, 0) = "Herr", "Herrn", infos(0, 0)), False
   agpreserve
   bookm "ha10a", zuh(infos(10, 0)), True
   agtabu
   Dim ii%, Index%
   For ii = 1 To UBound(infos, 2)
    If infos(4, ii) <> infos(4, 0) Then
     Index = ii
     Exit For
    End If
   Next ii
   bookm "ha1b", IIf(Index <> 0 And UBound(infos, 2) > 0, IIf(infos(0, Index) = "Herr", "Herrn", infos(0, Index)), "$1210$"), IIf(Index <> 0 And UBound(infos, 2) > 0, False, True)
   agpreserve
   bookm "ha10b", IIf(Index <> 0 And UBound(infos, 2) > 0, infos(10, Index), " "), IIf(Index <> 0 And UBound(infos, 2) > 0, False, True)
   ag.Append "</w:p>"
   ag.Append agabsa
   agtab2
   bookm "ha2a", zuh(infos(1, 0))
   agtabu ' Name:
   bookm "ha2b", zuh(IIf(Index <> 0 And UBound(infos, 2) > 0, infos(1, Index), "$1203$ $1202$ $1201$")), IIf(Index <> 0 And UBound(infos, 2) > 0, False, True)
   ag.Append "</w:p>"
   ag.Append agabsa
   agtab2
   bookm "ha3a", zuh(infos(2, 0))
   agtabu ' Straße:
   bookm "ha3b", zuh(IIf(Index <> 0 And UBound(infos, 2) > 0, infos(2, Index), "$1205$ $1206$")), IIf(Index <> 0 And UBound(infos, 2) > 0, False, True)
   ag.Append "</w:p>"
   ag.Append agabsa
   ag.AppVar Array("<w:pPr><w:framePr w:w=""340"" w:h=""550"" w:hSpace=""181"" w:wrap=""around"" w:vAnchor=""page"" w:hAnchor=""page"" w:x=""642"" w:y=""5671""/><w:tabs><w:tab w:val=""left"" w:pos=""5387""/></w:tabs><w:rPr><w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial""/><w:sz w:val=""16""/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial""/><w:sz w:val=""32""/></w:rPr><w:t>", zuh("-"), "</w:t></w:r></w:p>")
   ag.Append agabsa
   agtab2
   bookm "ha4a", zuh(infos(3, 0)), , True
   agtabu ' Ort:
   bookm "ha4b", zuh(IIf(Index <> 0 And UBound(infos, 2) > 0, infos(3, Index), "$1207$ $1208$")), IIf(Index <> 0 And UBound(infos, 2) > 0, False, True), IIf(Index <> 0 And UBound(infos, 2) > 0, False, True)
   ag.Append "</w:p>"
   ag.Append agabsa
   ag.Append doagtab & stil(, , 1, 2, , 1) & "</w:pPr>"
   ag.AppVar Array(bookmS("ha5a"), "<w:r>", stil(, , 1, 2, , 1), htxt(infos(4, 0)), "</w:r>", bookmE) ' FaxNr
   ag.AppVar Array("<w:r>", stil(, , 1, 2, , IIf(UBound(infos, 2) > 0, 0, 1)), "</w:r>")
   ag.AppVar Array(bookmS("ha5b"), "<w:r>", stil(, , 1, 2, , IIf(UBound(infos, 2) > 0, 0, 1)), htxt(IIf(Index <> 0 And UBound(infos, 2) > 0, infos(4, Index), "Fax $1216$:")), "</w:r>", bookmE) ' FaxNr
   ag.Append "</w:p>"
   Dim iru%
   For iru = 0 To 2
    ag.Append agabsa
    ag.Append "<w:pPr><w:tabs><w:tab w:val=""left"" w:pos=""6946""/><w:tab w:val=""left"" w:pos=""9214""/></w:tabs><w:ind w:right=""-1""/><w:rPr><w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial""/><w:b/><w:sz w:val=""24""/><w:u w:val=""single""/></w:rPr></w:pPr></w:p>"
   Next iru
   ag.Append agabsa
   ag.Append stil()
   bookm "ha6a", zuh(infos(5, 0)) & ","
   agpreserve
   ag.AppVar Array(bookmS("ha6b"), "<w:r>", stil(, , 1, 2, , IIf(UBound(infos, 2) > 0, 0, 1)), htxt(IIf(Index <> 0 And UBound(infos, 2) > 0, infos(5, Index), "$1213$")), "</w:r>", bookmE)
   agpreserve
   ag.Append "</w:p>"
   ag.Append agabsa
   ag.AppVar Array("<w:pPr><w:spacing w:before=""120""/>", stil(, , , 1), "</w:pPr>")
   ag.Append stiltxt("besten Dank für die freundliche Überweisung von ")
   ag.Append stiltxt(REPLACE$(GName, ", *", ", geb. ") & ", ", True)
   ag.Append stiltxt(dieder & " sich " & behDauerStr(Pat_ID, lbeh) & " bei uns vorstellte.")
   ag.Append "</w:p>"
   ag.Append agabsa
   ag.Append stil()
   ag.Append "</w:p>"
   ag.Append agabsa
   ag.Append doagtab("7938") & "<w:ind w:left=""1276"" w:hanging=""1276""/>" & stil(, , 1) & "</w:pPr>"
   ag.Append bookmS("Kompr") & stiltxt("Diagnosen", True, , True)
   ag.Append stiltxt(":" & Chr(9))
   ag.Append "</w:p>"
   ag.Append agabsa
   ag.Append doagtab("7938") & stil(, , 1) & "</w:pPr>"
   Dim dj%
   For dj = 0 To UBound(DiagTab)
    ag.Append IIf(dj > 0, "<w:br/>", "") & stiltxt(DiagTab(dj).Value)
   Next dj
   ag.Append bookmE
   If False Then
   End If
   ag.Append "</w:p>"
'  ag.Append agabsa
'  ag.AppVar Array("<w:pPr>", stil(, , 1), "<w:pPr/>")
  
'  ag.Append "<w:r>"
'  ag.Append "</w:r>"
   ag.Append "</w:body></w:document>"
   Print #51, zsuh(ag.Value)
'   Print #51, ConvertToUTF8(ag.Value)
   Close #51
   If False Then FSO.CopyFile vVerz & "h\word\document.xml", vVerz & "exp9\word\document.xml"
   oSh.rUn "powershell ""$vz=\""" & vVerz & "exp9\word\\\"";$dt=@($vz+\""endnotes.xml\"";$vz+\""footnotes.xml\"";$vz+\""document.xml\"";$vz+\""settings.xml\"");$anr=[string](get-content -path $dt[0]) -replace '.*w:rsidRDefault=\""([0-9A-F]*)\"".*','$1';$nrd=\""" & dzahl & "\"";$nr=[int](get-content -path $nrd)+1;set-content -path $nrd $nr;$nrs='{0:X8}' -f $nr;foreach ($dta in $dt){(get-content -path $dta) -replace $anr, $nrs|set-content -path $dta;};""", 0, True
  End If ' nichtherricht
  oSh.rUn "cmd /c """"c:\program files\7-zip\7z"" a -tzip -mm=deflate -mx9 -aoa -xr!*.swp """ & sverz & "\" & aname & "x"" " & vVerz & "exp9\*""", 0, True
' oSh.run "cmd /c """"c:\program files (x86)\microsoft office\root\office16\winword"" """ & sverz & aname & "x""""", 0, True
  oSh.rUn "cmd /c """"c:\program files\microsoft office\root\office16\winword"" """ & sverz & aname & "x""""", 0, True
 Else ' briefneu
'  ON Error Resume Next
  Call GetWord
   With Wapp
    .options.SmartCutPaste = False
    If WappBuild > 9 Then '.options("SmartParaSelection") = 0 '.Options.SmartParaSelection = False
     With .options
      .SmartParaSelection = 0
     End With
    End If
'   .Documents.Add Template:="c:\turbomed\vorlagen\AccessBriefa.dot" ' replace$(replace$(An1Pfad, "$\TurboMed\Dokumente\", PCDokPfad), "^", vns)
    Dim tpl$
neufestleg:
    Set dc = Nothing
    While dc Is Nothing
     tpl = getDokPfad("Vorlagen") & "\" & Vorlage ' xVerz & vorlagen\AccessBrief.dot
     On Error Resume Next
     Set dc = .documents.Add(Template:=tpl)
     On Error GoTo fehler
     If dc Is Nothing Then
      Dim erg$
      erg = MsgBox("Fehler beim Anlegen eines Winword-Dokumentes auf der Grundlage von " + tpl + ", welches aus den Turbomed Grundeinstellungen gelesen wurde! Erneut versuchen?", vbOKCancel, "Rückfrage")
      If erg = vbCancel Then Exit Sub
     End If
    Wend ' dc Is Nothing
    If nurLabor = 0 Then
     .ScreenUpdating = False
     .Visible = False
    End If ' nurlabor<>=
    With dc.bookmarks
'    SET dc = .ActiveDocument
    VorDat = GetVorDat(Pat_ID, obStumm)
    If nurLabor = 0 Then
On Error Resume Next
     Dim bzahl%
     bzahl = -1
     bzahl = .COUNT
On Error GoTo fehler
     If bzahl = -1 Then GoTo neufestleg
     Call BMRd(dc, !ha1a, IIf(infos(0, 0) = "Herr", "Herrn", infos(0, 0)), True)
     Call BMRd(dc, !ha2a, infos(1, 0)) ' Name
     Call BMRd(dc, !ha3a, infos(2, 0)) ' Straße
     Call BMRd(dc, !ha4a, infos(3, 0)) ' Ort
     Call BMRd(dc, !ha5a, infos(4, 0)) ' Faxnummer
     Call BMRd(dc, !ha6a, infos(5, 0)) ' Anrede
     Call BMRd(dc, !ha10a, infos(10, 0)) ' Überweiserinfos
'     Dim ii%, Index%
     For ii = 1 To UBound(infos, 2)
      If infos(4, ii) <> infos(4, 0) Then
       Index = ii
       Exit For
      End If
     Next ii
     If Index <> 0 And UBound(infos, 2) > 0 Then
      Call BMRd(dc, !ha1b, IIf(infos(0, Index) = "Herr", "Herrn", infos(0, Index)), True) '  IIF(raHacall bmrd(dc,!Geschlecht = "w", "Frau", "Herrn")) '+ vblf) + nzw
      Call BMRd(dc, !ha2b, infos(1, Index), True) ' "Dr.med. " + raHacall bmrd(dc,!VorName & " " & raHacall bmrd(dc,!NachName
      Call BMRd(dc, !ha3b, infos(2, Index), True) ' raHacall bmrd(dc,!Straße
      Call BMRd(dc, !ha4b, infos(3, Index), True) ' raHacall bmrd(dc,!Plz & " " & raHacall bmrd(dc,!Ort
      Call BMRd(dc, !ha5b, infos(4, Index)) ' Faxnummer
      If Not InStr(infos(10, 0), infos(10, Index)) And Not InStr(infos(10, Index), infos(10, 0)) Then
       Call BMRd(dc, !ha6b, REPLACE(infos(5, Index), "Sehr", " sehr") & ",", True) ' Anrede
      End If
      Call BMRd(dc, !ha10b, infos(10, Index)) ' Überweiserinfos
     End If ' Index <> 0 And UBound(infos, 2) > 0 Then
'weiter:
'    Next i
    !PVorn.Range = tit & VorNa
    !pnachn.Range = NachNa
    !PGeb.Range = Format$(GebDat, "d.M.YYYY")
    
'    Dim obAnredeSpez%
'    obAnredeSpez = 0
'    IF raHa.State = 1 THEN
'     IF Not (raHa.BOF OR raHa.EOF) THEN
'      obAnredeSpez = -1
'     END IF
'    END IF
'    IF obAnredeSpez THEN
'     SELECT CASE raHa!Überschrift
'      Case "L":  !haanr.Range = IIF(raHa!Geschlecht = "w", "Liebe", "Lieber") & " " & raHa!VorName
'      Case "H":  !haanr.Range = "Hallo " + raHa!VorName
'      Case Else: !haanr.Range = IIF(raHa!Geschlecht = "w", "Sehr geehrte Frau Kollegin", "Sehr geehrter Herr Kollege")
'     END SELECT
'    Else
''    IF !haanr.Range = "$1313$" THEN
'     !haanr.Range = "Sehr geehrte Frau Kollegin, sehr geehrter Herr Kollege"
'    END IF
    
    syscmd acSysCmdSetStatus, "Erstelle Brief für " & gesname & ": 2)a) Diagnosen prüfen ..."
    obkNeph = obKeineNephropathie(Pat_ID, obMakroAlb)
    Dim raFa As New ADODB.Recordset
    Dim lddat As Date
    myFrag raFa, "SELECT MAX(bhfb) lddat, MAX(fanf) fanf, MAX(ausgst) ausgst FROM `faelle` WHERE pat_id = " & Pat_ID
    If Not raFa.BOF Then
     If Not IsNull(raFa!lddat) Then
      lddat = raFa!lddat
     End If
     If raFa!Fanf > lddat Then lddat = raFa!Fanf
     If raFa!ausgst > lddat Then lddat = raFa!ausgst
    End If
    Dim iDiag As New ADODB.Recordset, licd$, ldiag$, lSicher$, iDDiagText$, iDICD$, iDDiagSich$
'    iDiag.Open "SELECT ICD, DiagText, DiagSicherheit FROM `diagnosen` d WHERE d.pat_id = " & Pat_id & " AND (d.obdauer <> 0 OR d.diagdatum > " & DatFor_k(lddat) & ") AND COALESCE(d.Dggel,0)=0 ORDER BY icd, DiagSicherheit", DBCn, adOpenDynamic, adLockReadOnly
    myFrag iDiag, "SELECT ICD, DiagText, DiagSicherheit FROM `diagnosen` d WHERE d.pat_id = " & Pat_ID & " AND (d.obdauer <> 0 OR d.diagdatum > " & DatFor_k(lddat) & ") ORDER BY icd, DiagSicherheit"
    If Not iDiag.BOF Then
'     iDiag.MoveFirst
     Do While Not iDiag.EOF
       iDDiagText = iDiag!DiagText
       iDICD = iDiag!ICD
       iDDiagSich = iDiag!DiagSicherheit
       Dim tonRunde%
'       MsgBox "Achtung doppelte Diagnose: " + nzw + _
              ldiag & " " & licd + nzw + _
              iDiag!DiagText & " " & iDiag!ICD
'      Call Shell("cmd /c echo ""Achtung doppelte Diagnose"" ", vbNormalNoFocus)
      If iDICD = licd And iDDiagSich = lSicher Then
       For tonRunde = 1 To 5
        Call Sound(WinDir + "\media\Windows XP-Standard.wav")
       Next tonRunde
       Call meld("Pat-ID " & Pat_ID & ", Achtung doppelte Diagnose: " & vbCrLf & ldiag & " " & licd & vbCrLf & iDDiagText & " " & iDICD, obStumm)
      End If ' iDICD = licd And iDDiagSicherheit = lSicher Then
      licd = iDICD
      ldiag = iDDiagText
      lSicher = iDDiagSich
      iDiag.Move 1
     Loop ' While Not iDiag.EOF
    End If ' Not iDiag.BOF Then
    syscmd acSysCmdSetStatus, "Erstelle Brief für " & gesname & ": 2)b) Diagnosen einfügen ..."
    !Diagnosen.Range = DString
    !pdieder.Range = dieder
    Dim behD As New CString
    behD = behDauerStr(Pat_ID, lbeh)
    !Zeitraum.Range = behD.Value
    If InStrB(behD, "am") <> 0 Then
     On Error Resume Next ' 16.10.14: entfällt, da umgestellt auf immer "besten Dank für die freundliche Überweisung"
     !Mitbehandlung.Range = "Untersuchung"
     On Error GoTo fehler
    End If
    syscmd acSysCmdSetStatus, "Erstelle Brief für " & gesname & ": 3) Anamnese ..."
    !Anamnese.Range = machwertString$(Pat_ID) '+ nzw + vblf
'    !Verlauf.Range = Verlauf(pat_id)
'END IF
    Dim Ereih%, TMn$, TZMn$, TMnErg$, krit$, Zerleg$(), Zerl2$(), Inhalt$
    Dim rlAb As New ADODB.Recordset, komm$, lauf%, KSpl$()
    syscmd acSysCmdSetStatus, "Erstelle Brief für " & gesname & ": 4) Weitere Daten ..."
'    rlAb.Open "SELECT DISTINCT quelle,Eingang,k.text kommentar FROM `laboryus` u INNER JOIN `laborybakt` b ON u.id = b.usid LEFT JOIN laboryhinw k ON b.kommid=k.id WHERE pat_id = " & pat_id & " AND NOT ISNULL(keimzahl) AND eingang > " & DatFor_k(VorDat), DBCn, adOpenDynamic, adLockReadOnly
'    rlAb.Open "SELECT quelle,Eingang, Kommentar FROM labor2bakt WHERE pat_id = " & Pat_id & " AND NOT ISNULL(keimzahl) AND eingang > " & DatFor_k(VorDat), DBCn, adOpenDynamic, adLockReadOnly
    myFrag rlAb, "SELECT quelle,Eingang, Kommentar FROM labor2bakt WHERE pat_id = " & Pat_ID & " AND NOT ISNULL(keimzahl) AND eingang > " & DatFor_k(VorDat), adOpenStatic
    Const bmUrin% = 7
    Const bmFuß% = 9
    Const bmEKG% = 13
    Const bmTaille% = 14
    Const bmLufu% = 16
    Const bmLZRR% = 17
    Dim bmTxt$(13 To 17)
    bmTxt(13) = "QRS"
    bmTxt(16) = "FEV"
    bmTxt(17) = "%Nacht"
    
    For Ereih = 1 To 19
'     IF Ereih = 5 THEN Stop
     TMn = "e" + CStr(Ereih)
     TZMn = "ez" + CStr(Ereih)
     krit = vNS
     On Error Resume Next
     Zerleg = Null
     Zerleg = Split(dc.bookmarks(TMn).Range, "[")
'     IF TMn = "e9" THEN Stop ' Ulucs
     On Error GoTo fehler
     If Not IsNull(Zerleg) Then
      For j = 0 To UBound(Zerleg)
       If InStrB(Zerleg(j), "]") > 0 Then
        Zerl2 = Split(Zerleg(j), "]")
        krit = krit + """" + Zerl2(0) + """" + ","
       End If
      Next j
      If InStr(krit, "notiz") Then
       krit = "'notiz','telef'," & artSpezBerat & "," & artSpezMA & ",'colo','coloauf','aug','augen','beruf','rauch','tv'"
      Else
       krit = Left$(krit, Len(krit) - 1) '+ """"
      End If
      Inhalt = eintraege(Pat_ID, krit, VorDat)
     End If
     Select Case Ereih
      Case bmTaille ' Bauchumfang / Taille
       Dim rTail As New ADODB.Recordset
       sql = "SELECT zeitpunkt AS zp, inhalt FROM `eintraege` WHERE pat_id = " & Pat_ID & " AND TRIM(art) = ""taille"" ORDER BY zeitpunkt"
       Set rTail = Nothing
'       rTail.Open sql, DBCn, adOpenDynamic, adLockReadOnly
       myFrag rTail, sql
       If rTail.BOF Then
       Else
        Inhalt = vNS
        Do While Not rTail.EOF
         Inhalt = Inhalt + CStr(DateValue(rTail!Zp)) + vbTab + rTail!Inhalt + nzw
         rTail.Move 1
        Loop
       End If ' rTail.BOF THEN
      Case bmUrin ' Urinuntersuchung
       Inhalt = REPLACE$(REPLACE$(Inhalt, "hts:", "Harnteststreifen:"), "hts.:", "Harnteststreifen:")
       If rlAb.State = 0 Then
        Dim source$
        source = rlAb.source
        Set rlAb = Nothing
        myFrag rlAb, source, adOpenStatic
       End If
       If Not rlAb.BOF Then
        Do While Not rlAb.EOF
         If Not IsNull(rlAb!Kommentar) Then
          KSpl = Split(rlAb!Kommentar, nzw)
          For lauf = 0 To UBound(KSpl)
           If InStrB(KSpl(lauf), ":") > 0 And InStrB(KSpl(lauf), "  ") > 0 Then
            KSpl(lauf) = REPLACE$(KSpl(lauf), " ", vNS)
           End If
           If Mid$(KSpl(lauf), 2, 4) = "Keim" Then If Right$(KSpl(lauf), 1) <> ":" Then KSpl(lauf) = KSpl(lauf) + ":"
          Next lauf
          For lauf = 0 To UBound(KSpl) - 1
           If Right$(KSpl(lauf), 1) = ":" Then
             KSpl(lauf) = KSpl(lauf) + IIf(Asc(Left$(KSpl(lauf + 1), 1)) < 32, Mid$(KSpl(lauf + 1), 2), KSpl(lauf + 1))
             KSpl(lauf + 1) = vNS
           End If
          Next lauf
          komm = vNS
          For lauf = 0 To UBound(KSpl)
           KSpl(lauf) = REPLACE$(KSpl(lauf), ":", ":" + vbTab)
           If KSpl(lauf) <> vNS Then
             komm = komm + nzw + KSpl(lauf)
           End If
          Next lauf
          If komm <> vNS Then
           Inhalt = Inhalt + nzw + rlAb!Quelle + " (" & rlAb!Eingang & "):" + vbTab + komm + nzw
          End If
         End If ' NOT ISNULL(rlAb!Kommentar) Then
         rlAb.Move 1
        Loop ' While Not rlAb.EOF
       End If ' Not rLab.BOF THEN
      Case bmFuß
'       Inhalt = replace$(Inhalt, nzw, ", ")
'       sql = "SELECT zeitpunkt AS zp, inhalt FROM `eintraege` WHERE pat_id = " & pat_id & " AND TRIM(art) IN ('usdm','usdm1','fuß','fuss') AND zeitpunkt > " & DatFor_k(VorDat) & " ORDER BY zeitpunkt"
''       SET rTail = Dtb.OpenRecordset(sql)
'       rTail.Open sql, DBCn, adOpenDynamic, adLockReadOnly
'       IF rTail.BOF THEN
'       Else
'        Inhalt = vNS
'        Do While Not rTail.EOF
'         Dim InhVorl$
'         pos = InStr(rTail!Inhalt, "aktuellen Blutdruck")
'         IF pos > 0 THEN InhVorl = LEFT(rTail!Inhalt, pos - 1) ELSE InhVorl = rTail!Inhalt
'         Inhalt = Inhalt + CStr(DateValue(rTail!Zp)) + vbTab + InhVorl + nzw
'         rTail.Move 1
'        Loop
'        Inhalt = LEFT(Inhalt, Len(Inhalt) - 1)
'       END IF
      Case bmEKG, bmLufu, bmLZRR
'       sql = "SELECT FormInhaltNeu.ZeitPunkt, `forminhaltfeld`.Feld, `forminhaltfeldinh`.FeldInh, FormInhaltNeu.AbsPos " + _
             "FROM FormInhaltForm_Abk INNER JOIN (`forminhaltfeldinh` INNER JOIN (`forminhaltfeld` INNER JOIN FormInhaltNeu ON `forminhaltfeld`.FeldVW = FormInhaltNeu.FeldVW) ON `forminhaltfeldinh`.FeldInhVW = FormInhaltNeu.FeldInhVW) ON FormInhaltForm_Abk.Form_AbkVW = FormInhaltNeu.Form_AbkVW " + _
             "WHERE (Feld =""Zeile"" OR ISNULL(Feld)) AND FeldInh NOT LIKE ""-*"" AND FeldInh NOT LIKE ""Ruhe-EKG Messwerte*"" AND Form_Abk=""GDT"" AND pat_id = " & pat_id
'       sql = " SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, FeldVW, FeldInhVW FROM ((`forminhfeld` LEFT JOIN `forminhkopf` ON `forminhfeld`.foid = `forminhkopf`.foid) LEFT JOIN `formulare` ON `formulare`.formid = `forminhkopf`.form_id) LEFT JOIN forminhaltform_abk ON forminhaltform_abk.form_abkvw = `formulare`.form_abkvw WHERE (Feld =""Zeile"" OR ISNULL(Feld)) AND FeldInh NOT LIKE ""-*"" AND FeldInh NOT LIKE ""Ruhe-EKG Messwerte*"" AND Form_Abk=""GDT"" AND pat_id = " & pat_id
'       sql = "SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, Form_Abk FROM (((`forminhfeld` LEFT JOIN `forminhkopf` ON `forminhfeld`.foid = `forminhkopf`.foid) LEFT JOIN `formulare` ON `formulare`.formid = `forminhkopf`.form_id) LEFT JOIN `forminhaltfeld` ON `forminhfeld`.feldvw = `forminhaltfeld`.feldvw) LEFT JOIN `forminhaltfeldinh` ON `forminhfeld`.feldinhvw = `forminhaltfeldinh`.feldinhvw WHERE (Feld =""Zeile"" OR ISNULL(Feld)) AND FeldInh NOT LIKE '-%' AND Form_Abk IN ('GDT','EKG','Lufu','LZRR') AND pat_id = " & pat_id & " AND Zeitpunkt > " & DatFor_k(vordat)
'       vordat = CDate("1.1.1900")
       Dim rEkg As New ADODB.Recordset
       sql = "SELECT rfi,Pat_ID,form_abk,Form_ID,fif.foid AS ffoid,ZeitPunkt,Nr,FeldNr,Feld,FeldInh FROM (((`forminhfeld` fif LEFT JOIN `forminhkopf` fik ON fif.foid = fik.foid) LEFT JOIN `formulare` ON `formulare`.formid = fik.form_id) LEFT JOIN `forminhaltfeld` fihf ON fif.feldvw = fihf.feldvw) LEFT JOIN `forminhaltfeldinh` fihfi ON fif.feldinhvw = fihfi.feldinhvw LEFT JOIN (SELECT rfif.foid,feldinh AS rfi,pat_id AS rpat_id FROM `forminhaltfeldinh` rfihfi LEFT JOIN `forminhfeld` rfif ON rfihfi.feldinhvw = rfif.feldinhvw LEFT JOIN `forminhkopf` rfik ON rfif.foid = rfik.foid WHERE rfihfi.feldinh LIKE '" & bmTxt(Ereih) & "%' AND pat_id = " & Pat_ID & " AND Zeitpunkt > " & DatFor_k(VorDat) & " GROUP BY foid) AS rfoid ON fik.foid = rfoid.foid WHERE (Feld ='Zeile' OR ISNULL(Feld)) AND FeldInh NOT LIKE '-%' AND Form_Abk IN ('GDT','EKG','LZRR','Lufu') AND pat_id = " & Pat_ID & " AND Zeitpunkt > " & DatFor_k(VorDat) & " AND NOT ISNULL(rfi) ORDER BY pat_id,form_abk,fif.foid,zeitpunkt,nr"
       Set rEkg = Nothing
'       rEkg.Open sql, DBCn, adOpenDynamic, adLockReadOnly
       myFrag rEkg, sql
       Dim altfoid&, FeldInh$, nfi$
       Inhalt = vNS
       altfoid = -1
       Do While Not rEkg.EOF
        FeldInh = REPLACE$(REPLACE$(REPLACE$(REPLACE$(rEkg!FeldInh, "|", ","), " ¦", " °"), " ,", ","), ",,", ",")
        Do
         nfi = REPLACE$(FeldInh, "  ", " ")
         If nfi = FeldInh Then
          Exit Do
         Else
          FeldInh = nfi
         End If
        Loop
        FeldInh = Left$(FeldInh, Len(FeldInh) - IIf(Right$(FeldInh, 1) = ",", 1, 0))
'          Inhalt = Inhalt & ", " & FeldInh
        If FeldInh Like "ST_*" Or FeldInh Like "Untersuchung vom*" Or FeldInh = "Ruhe-EKG Messwerte" Then
        Else
         Inhalt = Inhalt & IIf(LenB(Inhalt) = 0, vNS, nzw) & IIf(rEkg!ffoid <> altfoid, Format$(rEkg!Zeitpunkt, "dd/mm/yy"), "") + vbTab + FeldInh
         altfoid = rEkg!ffoid
        End If
        rEkg.Move 1
       Loop
     End Select
'      IF Ereih = 4 THEN Stop
     If LenB(Inhalt) = 0 Then
'      IF instrb(dc.Bookmarks(TMn).Range, "$2004") > 0 THEN
       On Error Resume Next
       dc.Range(dc.bookmarks(TMn).Range.Start, dc.bookmarks(TMn).Range.END + 1) = Inhalt
       dc.bookmarks(TZMn).Range = vNS
       On Error GoTo fehler
'      END IF
     Else
       On Error GoTo fehler
'      IF instrb(dc.Bookmarks(TMn).Range, "$2004") > 0 THEN
       If Ereih = 4 Then
        If InStr(Inhalt, Chr(10)) = Len(Inhalt) Then
         dc.Range(dc.bookmarks("ez4").Range.Start, dc.bookmarks("ez4").Range.END - 1) = "Gewicht:" ' statt Gewichtsverlauf
        End If
       End If
       dc.bookmarks(TMn).Range = Inhalt
'      END IF
     End If
    Next Ereih
    Dim rlAbsrc$
    rlAbsrc = rlAb.source
    If rlAb.State = 0 Then
     Set rlAb = Nothing
     rlAb.Open rlAbsrc, DBCn, adOpenStatic, adLockReadOnly
    End If
    If rlAb.State <> 0 Then
     If Not rlAb.BOF Then
      Dim aktr As Range
      Dim aktRa
      Set aktRa = dc.bookmarks("Urin").Range
'     With aktra.Find
'     .ClearFormatting
'     .text = nzw + nzw '"^p^p"
'     .Execute
'     END With
'     IF aktra.Find.Found THEN
'      SET aktra = dc.Range(aktra.Start + 2, dc.Bookmarks("Urin").End)
      pos = InStr(aktRa, nzw + nzw)
      If pos > 0 Then
       Set aktRa = dc.Range(aktRa.Start + pos + 1, dc.bookmarks("Urin").END)
       aktRa.ParagraphFormat.TabStops.ClearAll
       aktRa.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(4)
       aktRa.ParagraphFormat.FirstLineIndent = CentimetersToPoints(0)
       aktRa.ParagraphFormat.LeftIndent = 0
      End If ' pos > 0 THEN
     End If ' Not rlAb.BOF THEN
    End If
   End If ' NurLabor
  End With ' dc.bookmarks
    syscmd acSysCmdSetStatus, "Erstelle Brief für " & gesname & ": 5) Labor..."
    Call LaborIns1(dc, Pat_ID, nurLabor)
    If nurLabor = 0 Then
    syscmd acSysCmdSetStatus, "Erstelle Brief für " & gesname & ": 6) Medikation..."
    Wapp.Visible = True
    Call letzteMed(dc, Pat_ID)
    syscmd acSysCmdSetStatus, "Erstelle Brief für " & gesname & ": 7) Epikrise..."
    Call Epikrise(dc, Pat_ID, VorDat, lddat, obStumm)
'      dc.bookmarks("DMP").Range = vns
      If Now - lddat > 90 Then
       If Now - lddat > 365 Then
        dc.bookmarks("DMPText").Range = "Ich bedauere die beträchtliche Verspätung des Arztbriefes." & vbCrLf
       Else
        dc.bookmarks("DMPText").Range = "Ich bedauere die Verspätung des Arztbriefes." & vbCrLf
       End If
      Else
       dc.bookmarks("DMPText").Range = vNS
      End If
'      dc.bookmarks("DMPText").Range = vns
    On Error GoTo fehler
' bei Wiederholungsbriefen
    syscmd acSysCmdSetStatus, "Erstelle Brief für " & gesname & ": 9) Prüfen von Vorberichten..."
    Call WiederHolungsBrief(NachNa, VorDat, Wapp, dc)
    syscmd acSysCmdSetStatus, "Erstelle Brief für " & gesname & ": 10) Korrekturen..."
    Call dc.Range.Find.Execute("Markt^t", , , , , , , , , "Markt Indersdorf^t", wdReplaceAll)
    Call dc.Range.Find.Execute("IDA", True, , , , , , , , "Insulindosisanpassung", wdReplaceAll)
    Call dc.Range.Find.Execute("ID-Anpassung", True, , , , , , , , "Insulindosisanpassung", wdReplaceAll)
    Call dc.Range.Find.Execute("Hypo`s", True, , , , , , , , "Hypoglykämien", wdReplaceAll)
    Call dc.Range.Find.Execute("GPD", True, , , , , , , , "Gesundheitspaß Diabetes", wdReplaceAll)
    Call dc.Range.Find.Execute("DFS", True, , , , , , , , "Diabetisches Fußsyndrom", wdReplaceAll)
    Call dc.Range.Find.Execute("bettunge,", True, , , , , , , , "bettungen,", wdReplaceAll)
    Call dc.Range.Find.Execute("bettunge.", True, , , , , , , , "bettungen.", wdReplaceAll)
    Call dc.Range.Find.Execute("$2032$", True, , , , , , , , vNS, wdReplaceAll)
    Call dc.Range.Find.Execute("n.u.", True, , , , , , , , "nicht untersucht", wdReplaceAll)
    Dim suchr
    On Error Resume Next
    Set suchr = dc.bookmarks("Zeitraum")
    If Err.Number <> 0 Then Err.Clear: Set suchr = dc.bookmarks("Anamnese")
    If Err.Number = 0 Then
     Call dc.Range(suchr.Range.Start, dc.Range.END).Find.Execute("^p^p^p^p", True, , , , , , , , "^p^p", wdReplaceAll)
     Call dc.Range(suchr.Range.Start, dc.Range.END).Find.Execute("^p^p^p^p", True, , , , , , , , "^p^p", wdReplaceAll)
    End If
    On Error GoTo fehler
    Set suchr = dc.Range
    Call suchr.Find.Execute("Verlaufsnotizen")
    dc.Range(suchr.Start, suchr.Start).Select
   End If ' NurLabor
   If Not obStumm Then .Visible = True
   .Application.WindowState = wdWindowStateMaximize
   .ScreenUpdating = True
   If Not obStumm Then .Activate
   syscmd acSysCmdSetStatus, "Erstelle Brief für " & gesname & ": 11) Speichern..."
   Call ZwischenSpeichern(sverz, NachNa, VorNa, Pat_ID, infos) ' infos(4, 0), infos(15, 0)) ' Faxnr, Mail
  End With ' wapp
  On Error Resume Next
  If obStumm Then
   Call Wapp.Quit(0)  '  = wdDoNotSaveChanges
  Else ' obStumm
   Call snie
  End If ' obStumm
 End If ' briefneu
 Lese.Ausgeb "Fertig mit Arztbrief für " & pid & "(" & gesname & ")", True, True
 syscmd acSysCmdClearStatus
 Exit Sub
fehler:
 Dim FNr&, FLastDLLError&, FSource$, FDescr$
 FNr = Err.Number
 FLastDLLError = Err.LastDllError
 FSource = Err.source
 FDescr = Err.Description
 PiepKurz
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(FNr) + vbCrLf + "LastDLLError: " + CStr(FLastDLLError) + vbCrLf + "Source: " + IIf(IsNull(FSource), vNS, CStr(FSource)) + vbCrLf + "Description: " + FDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in tubriefStandalone/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' tubriefStandalone(frm AS Form)

Public Function zuh$(ByRef q$)
 Dim i&, a%, b$
 zuh = ""
 For i = 1 To Len(q)
  b = Mid$(q, i, 1)
  a = Asc(b)
  If a < 128 Then
   zuh = zuh & b
  Else ' a < 128
   zuh = zuh & "&#" & a & ";"
  End If ' a < 128
 Next i
End Function ' zuh

#If zutesten Then
Function leseu()
 Const Datei$ = "exp9\word\document.xml"
 Dim zeile$
' Lese.ProgStart
 Open vVerz & Datei For Input As #52
 While Not EOF(52)
  Line Input #52, zeile
  If InStr(zeile, "besten") <> 0 Then
   Debug.Print zeile
  End If
 Wend
' Lese.ProgEnde
 Close #52
End Function ' leseu()
#End If

' in tubriefStanndalone
' Ersetzt Bookmark mit Text, ohne es zu löschen
Function BMRd(dc As Object, Bm, Text$, Optional obSichtbar%)
 Dim Start&, obEnde&, neulen&, bmname$
 On Error GoTo fehler
 With Bm
  Start = .Start
  obEnde = .END
 End With
 neulen = Len(Text)
 bmname = Bm.name
 Bm.Range = Text
 Call dc.bookmarks.Add(bmname, dc.Range(Start, Start + neulen))
 If obSichtbar Then dc.bookmarks(bmname).Range.Font.Hidden = False
 Exit Function
fehler:
  Dim FNr&, FLastDLLError&, FSource$, FDescr$
  FNr = Err.Number
  FLastDLLError = Err.LastDllError
  FSource = Err.source
  FDescr = Err.Description
  PiepKurz
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(FNr) + vbCrLf + "LastDLLError: " + CStr(FLastDLLError) + vbCrLf + "Source: " + IIf(IsNull(FSource), vNS, CStr(FSource)) + vbCrLf + "Description: " + FDescr, vbAbortRetryIgnore, "Aufgefangener Fehler in BMRd/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' bmrd

Function fobHAimDMP%(HANr$)
 Dim rKVA As New ADODB.Recordset
 Dim obDMP2%, obDMP1%
 On Error GoTo fehler
' Call KVÄVorb
' Call acon(HaT)
 If LenB(DBCn) = 0 Or DBCn = "" Then Call acon(quelleT)
' rKVA.Open "SELECT -dmpt1 AS j_dmpt1, -dmpt2 AS j_dmpt2, hae.* FROM `kvaerzte`.`hae` WHERE kvnr = '" & Left$(HANr, 2) & "/" & Right$(HANr, 5) & "'", DBCn, adOpenDynamic, adLockReadOnly ' HAECn
 myFrag rKVA, "SELECT -dmpt1 AS j_dmpt1, -dmpt2 AS j_dmpt2, hae.* FROM `kvaerzte`.`hae` WHERE kvnr = '" & Left$(HANr, 2) & "/" & Right$(HANr, 5) & "'"
 If Not rKVA.BOF Then
  Do While Not rKVA.EOF
   If rKVA!j_dmpt2 <> 0 Then obDMP2 = True
   If rKVA!j_dmpt1 <> 0 Then obDMP1 = True
   rKVA.Move 1
  Loop
 End If
 fobHAimDMP = obDMP2
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in fobHAimDMP/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' fobHAimDMP(HANr$)

#If mitab Then
' in do_Form_Current_Anbog
Sub HAlokal(rHa As ADODB.Recordset, KVNr$, Optional NaN, Optional VoN)
#If False Then
' in der örtlichen Datei nicht gefundenen Hausarzt aus der systemischen übertragen
  Dim rKv As New ADODB.Recordset
  On Error GoTo fehler
       Set rKv = TabÖff("HAE")
       If Not IsMissing(NaN) And Not IsMissing(VoN) Then
        myFrag rKv, "SELECT * FROM `hae` WHERE kvnr = '" & Left$(KVNr, 2) & "/" & Right$(KVNr, 5) & "' AND nachname = '" & NaN & "' AND vorname = '" & VoN & "'", adOpenStatic, HAECn, adLockReadOnly
       End If
       If rKv Is Nothing Or rKv.BOF Then
        Set rKv = Nothing
        myFrag rKv, "SELECT * FROM `hae` WHERE kvnr = '" & Left$(KVNr, 2) & "/" & Right$(KVNr, 5) & "' AND nachname = '" & NaN & "'", adOpenStatic, HAECn, adLockReadOnly
       End If
       If rKv Is Nothing Or rKv.BOF Then
        Set rKv = Nothing
        myFrag rKv, "SELECT * FROM `hae` WHERE kvnr = '" & Left$(KVNr, 2) & "/" & Right$(KVNr, 5) & "'", adOpenStatic, HAECn, adLockReadOnly
       End If
       If Not rKv.BOF Then
        rHa.AddNew
        rHa!Nachname = rKv!Nachname
        rHa!Vorname = rKv!Vorname
        rHa!KVNr = Left$(REPLACE$(rKv!KVNr, "/", vNS), 7)
        rHa!anschrift = rKv!Straße & ", " & rKv!plz & " " & rKv!ort
        rHa!plz = rKv!plz
        rHa!ort = rKv!ort
        rHa!Straße = rKv!Straße
        rHa!telefon = rKv!tel1
        rHa!telefax = rKv!fax1
        rHa!e_mail = rKv!email
        rHa!zulassungsgebiet = rKv!zulg
        rHa!arzttyp = rKv!arzttyp
        rHa![Gemeinschaftspraxis mit] = rKv!gemmit
        rHa!Bemerkung = rKv!beme
        rHa!geschlecht = IIf(rKv!anrede = "Frau", "w", IIf(rKv!anrede = "Herr", "m", vNS))
        rHa.Update
'        rHa.Bookmark = rHa.LastModified
        Call do_haakt(rHa)
       End If
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in HAlokal/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
#End If
End Sub ' HAlokal
#End If

' obFälle falsch geordnet:
' SELECT * FROM (SELECT i.pat_id ipat_id, i.fid ifid, i.bhfb ibhfb, f2.bhfb f2bhfb, IF(f2.bhfb > i.bhfb,1,0) ob FROM (SELECT pat_id, fid, bhfb, (SELECT MAX(fid) FROM `faelle` WHERE pat_id = f1.pat_id AND fid < f1.fid) mfid FROM `faelle` f1) i LEFT JOIN `faelle` f2 ON f2.fid = i.mfid) i WHERE ob <> 0;
' => sind sie nicht
' Aufruf in getHausarzt1
Function Üwrd$(rFa() As Faelle, Üw1$(), rKv() As kvnrue, Optional auchwir% = 0)
' die Fallzeile mit dem jüngsten Behandlungsfallbeginn; rKV() AS kvnrue
' zunächst für das Anamneseblatt, später auch für den Brief
  Const maxFZ% = 50 ' Maximalzahl der verschiedenen Fälle
'  Dim rsNa As New ADODB.Recordset
'  Dim raFa As New ADODB.Recordset, raFaEOF%
  Dim raFaEOF%
  Dim stand%, runde%, Feld$, i%, obalt%, bhb#(), obneueFallZeit%, j&
  Dim iii%
  Dim Üw1ini% ' üw1 wurde initialisiert
vonvorne:
  ReDim bhb(0)
  On Error GoTo fehler
  stand = -1
'  Üw2 = vns
    
'  raFa.Open "SELECT * FROM `faelle` WHERE pat_id = " & Pat_id & " ORDER BY bhfb DESC", DBCn, adOpenDynamic, adLockReadOnly
'  SET rFa = TabÖff("faelle", "AktF") ' pat_id aufsteigend, BHFB absteigend
'  rFa.Seek "=", Pat_id
'  raFaEOF = raFa.EOF
'  IF Not raFa.EOF THEN
' Fälle werden nach zeitlicher Relevanz durchsucht
' Wenn ein neuer Überweiser auftaucht, wird Üw1 um 1 auf (stand-1) erweitert und zuletzt der neue Überweiser eingetragen
' wenn ein alter Überweiser, der zuvor evtl. nur aus AndÜW bekannt war, jetzt aus Übwv bekannt wird, so werden Name und Vorname ergänzt
   j = UBound(rFa)
   If j < 1 Then raFaEOF = True Else raFaEOF = 0
   Do
' Zahl der verschiedenen Fälle bestimmen, ggf Abbruch
    For runde = 0 To UBound(bhb)
     obneueFallZeit = True
     If bhb(runde) = 0 Then ' nur am Anfang
      obneueFallZeit = False
      bhb(runde) = rFa(j).BhFB
      Exit For
     ElseIf bhb(runde) = rFa(j).BhFB Then
      obneueFallZeit = False
      Exit For
     End If
    Next runde
    If obneueFallZeit Then
     ReDim Preserve bhb(UBound(bhb) + 1)
     bhb(UBound(bhb)) = rFa(j).BhFB
    End If
    If UBound(bhb) = maxFZ Then Exit Do
    
    For runde = 1 To 2 ' 19.4.11
     Select Case runde
      Case 1: Feld = rFa(j).Übwr
      Case 2: Feld = rFa(j).AndÜw
'      Case 2: Feld = Left$(IIF(rFa(j).ÜbWVBSNR = vNS, IIF(rFa(j).ÜbWVKVNR = vNS, rFa(j).AndÜw, rFa(j).ÜbWVKVNR), rFa(j).ÜbWVBSNR), 7) 'IF(übwvbsnr='',IF(übwvkvnr = '', andüw, übwvkvnr),übwvbsnr) üw
     End Select
     If Not IsNull(Feld) Then
      If Feld <> "0000000" And Feld <> vNS Then
       obalt = 0
       For i = 0 To stand
        If Üw1(0, i) = Feld Then
         If runde = 1 Then
          Üw1(1, i) = rFa(j).ÜWNaN
          Üw1(2, i) = rFa(j).ÜWVor
          Üw1(4, i) = rFa(j).ÜbWVKVNR
         End If
         obalt = True
         Exit For
        End If
       Next i
       If Not obalt And Feld <> KVNr And Feld <> BSNR And Feld <> "889690003" And Feld <> "933284903" Then
        stand = stand + 1
        ReDim Preserve Üw1(4, stand) ' 0 = KV-Nr, 1 = Nachname, 2 = Vorname, 3 = Position
        Üw1ini = True
        Üw1(0, stand) = Feld
        If InStrB(Üw1(0, stand), "TM#") <> 0 Then
         For iii = Len(Üw1(0, stand)) To 0 Step -1
          If Not IsNumeric(Mid$(Üw1(0, stand), iii, 1)) Then
           Üw1(0, stand) = Mid$(Üw1(0, stand), iii + 1)
'           IF Len(Üw1(0, stand)) > 7 THEN Üw1(0, stand) = Left$(Üw1(0, stand), 7)
           Exit For
          End If
         Next iii
        End If
        Üw1(3, stand) = "Üw " & Left$(rFa(j).Quartal, 1) & Right$(rFa(j).Quartal, 2)
        If runde = 1 Then
         Üw1(1, stand) = rFa(j).ÜWNaN
         Üw1(2, stand) = rFa(j).ÜWVor
         Üw1(4, stand) = rFa(j).ÜbWVKVNR
        End If
       End If ' not obalt
       Exit For
      End If ' Feld <> "0000000" AND Feld <> vNS THEN
     End If ' NOT ISNULL(Feld) THEN
    Next runde
'    IF stand >= maxÜwZ - 1 THEN Exit Do
    j = j - 1
    If j <= 0 Then Exit Do
'    IF raFa!Pat_id <> Pat_id THEN Exit Do
   Loop
 
'  SET rsNa = TabÖff("Namen", "pat_ID")
'  rsNa.Seek "=", Pat_id
'  myFrag rsNa, "SELECT k.kvnr FROM `namen` n LEFT JOIN kvnrue k ON n.pat_id = k.pat_id WHERE n.pat_id = " & Pat_id & " ORDER BY k.lfdnr"
  Dim obrKKvNr%
  obrKKvNr = 0
  Dim rK As New ADODB.Recordset
  If UBound(rFa) > 0 Then
   Dim obrKvzugew%
   If (0 / 1) + (Not Not rKv) = 0 Then
   Else
    For i = 0 To UBound(rKv)
     If rKv(i).KVNr <> "" Then
      obrKvzugew = True
      Exit For
     End If
    Next i
   End If ' (0 / 1) + (Not Not rKv) = 0 Then Else
   If (0 / 1) + (Not Not rKv) = 0 Or obrKvzugew = 0 Then
    myFrag rK, "SELECT kvnr FROM `kvnrue` WHERE pat_id = " & rFa(1).Pat_ID & IIf(auchwir, "", " AND kvnr NOT IN ('','" & KVNr & "','" & BSNR & "','889690003','933284903')") & " ORDER BY lfdnr"
    If Not rK.BOF Then
     Do While Not rK.EOF()
       obalt = 0
       obrKKvNr = True
       For i = 0 To stand
'        IF Left$(Üw1(0, i), 7) = rK!KVNr THEN
        If Üw1(4, i) = rK!KVNr Then
         Üw1(3, i) = Üw1(3, i) & ", HA"
         obalt = True
         Exit For
        End If
       Next i
       If Not obalt Then
        stand = stand + 1
        ReDim Preserve Üw1(4, stand) ' 0 = KV-Nr, 1 = Nachname, 2 = Vorname, 3 = Position
        Üw1ini = True
        Üw1(0, stand) = rK!KVNr
        Üw1(4, stand) = rK!KVNr
        Üw1(3, stand) = "HA"
       End If ' Not obalt THEN
       rK.MoveNext
     Loop
    End If ' Not rK.BOF THEN
   Else ' isMissing(rKV)
    Dim ii&
    For ii = 1 To UBound(rKv)
       obalt = 0
       obrKKvNr = True
       For i = 0 To stand
'        IF Left$(Üw1(0, i), 7) = rKv(ii).KVNr THEN
        If Üw1(4, i) = rKv(ii).KVNr Then
         Üw1(3, i) = Üw1(3, i) & ", HA"
         obalt = True
         Exit For
        End If
       Next i
       If Not obalt Then
        stand = stand + 1
        ReDim Preserve Üw1(4, stand) ' 0 = KV-Nr, 1 = Nachname, 2 = Vorname, 3 = Position
        Üw1ini = True
        Üw1(0, stand) = rKv(ii).KVNr
        Üw1(4, stand) = rKv(ii).KVNr
        Üw1(3, stand) = "HA"
       End If ' Not obalt THEN
    Next ii
   End If ' isMissing(rKV)
  End If ' UBound(rFa) > 0 THEN
  
  If Not raFaEOF Or obrKKvNr Then
'   ON Error Resume Next ' Index außerhalb des Bereichs
   If Üw1ini Then
    Üwrd = Üw1(0, 0)
    If LenB(Üwrd) = 0 Then Üwrd = "-"
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Üwrd/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Üwrd

Function Üw12$(Pat_ID&, Üw1$()) ' Saubere Funktion zum Ermitteln der Überweisernummern aus `faelle` sowie des Hausarztes aus `namen`
' zunächst für das Anamneseblatt, später auch für den Brief
  Const maxFZ% = 50 ' Maximalzahl der verschiedenen Fälle
  Dim rsNa As New ADODB.Recordset
  Dim raFa As New ADODB.Recordset, raFaEOF%
  Dim stand%, runde%, Feld$, i%, obalt%, bhb#(), obneueFallZeit%
  Dim iii%
  Dim Üw1ini% ' üw1 wurde initialisiert
  ReDim bhb(0)
  On Error GoTo fehler
  stand = -1
'  Üw2 = vns
    
'  raFa.Open "SELECT * FROM `faelle` WHERE pat_id = " & Pat_id & " ORDER BY bhfb DESC", DBCn, adOpenDynamic, adLockReadOnly
  myFrag raFa, "SELECT * FROM `faelle` WHERE pat_id = " & Pat_ID & " ORDER BY bhfb DESC"
'  SET rFa = TabÖff("faelle", "AktF") ' pat_id aufsteigend, BHFB absteigend
'  rFa.Seek "=", Pat_id
  raFaEOF = raFa.EOF
  If Not raFa.EOF Then
' Fälle werden nach zeitlicher Relevanz durchsucht
' Wenn ein neuer Überweiser auftaucht, wird Üw1 um 1 auf (stand-1) erweitert und zuletzt der neue Überweiser eingetragen
' wenn ein alter Überweiser, der zuvor evtl. nur aus AndÜW bekannt war, jetzt aus Übwv bekannt wird, so werden Name und Vorname ergänzt
   Do
' Zahl der verschiedenen Fälle bestimmen, ggf Abbruch
    For runde = 0 To UBound(bhb)
     obneueFallZeit = True
     If bhb(runde) = 0 Then ' nur am Anfang
      obneueFallZeit = False
      bhb(runde) = raFa!BhFB
      Exit For
     ElseIf bhb(runde) = raFa!BhFB Then
      obneueFallZeit = False
      Exit For
     End If
    Next runde
    If obneueFallZeit Then
     ReDim Preserve bhb(UBound(bhb) + 1)
     bhb(UBound(bhb)) = raFa!BhFB
    End If
    If UBound(bhb) = maxFZ Then Exit Do
    
    Dim Fertig%
    Fertig = 0
    For runde = 2 To 2
     Select Case runde
      Case 1: Feld = "AndÜw"
      Case 2: Feld = "Übwr"
'      Case 2: Feld = "Übwvbsnr" ' IF(übwvbsnr='',IF(übwvkvnr = '', andüw, übwvkvnr),übwvbsnr) üw
'      Case 3: Feld = "Übwvkvnr"
     End Select
     If Not IsNull(raFa.Fields(Feld)) Then
      If raFa.Fields(Feld) <> "0000000" And LenB(raFa.Fields(Feld)) <> 0 And Not Fertig Then
       If runde = 2 Then Fertig = True
       obalt = 0
       For i = 0 To stand
        If Üw1(0, i) = raFa.Fields(Feld) Then
         If runde > 1 Then
          Üw1(1, i) = raFa!ÜWNaN
          Üw1(2, i) = raFa!ÜWVor
          Üw1(4, i) = raFa!ÜbWVKVNR
         End If
         obalt = True
         Exit For
        End If
       Next i
       If Not obalt Then
        stand = stand + 1
        ReDim Preserve Üw1(4, stand) ' 0 = KV-Nr, 1 = Nachname, 2 = Vorname, 3 = Position
        Üw1ini = True
        Üw1(0, stand) = raFa.Fields(Feld)
        If InStrB(Üw1(0, stand), "TM#") <> 0 Then
         For iii = Len(Üw1(0, stand)) To 0 Step -1
          If Not IsNumeric(Mid$(Üw1(0, stand), iii, 1)) Then
           Üw1(0, stand) = Mid$(Üw1(0, stand), iii + 1)
           If Len(Üw1(0, stand)) > 7 Then Üw1(0, stand) = Left$(Üw1(0, stand), 7)
           Exit For
          End If
         Next iii
        End If
        Üw1(3, stand) = "Üw " & Left$(raFa!Quartal, 1) & Right$(raFa!Quartal, 2)
        If runde > 1 Then
         Üw1(1, stand) = raFa!ÜWNaN
         Üw1(2, stand) = raFa!ÜWVor
         Üw1(4, stand) = raFa!ÜbWVKVNR
        End If
       End If ' not obalt
      End If
     End If
    Next runde
'    IF stand >= maxÜwZ - 1 THEN Exit Do
    raFa.MoveNext
    If raFa.EOF Then Exit Do
'    IF raFa!Pat_id <> Pat_id THEN Exit Do
   Loop
  Else
   Exit Function
  End If
  
'  SET rsNa = TabÖff("Namen", "pat_ID")
'  rsNa.Seek "=", Pat_id
  myFrag rsNa, "SELECT k.kvnr FROM `namen` n LEFT JOIN `kvnrue` k ON n.pat_id = k.pat_id WHERE n.pat_id = " & Pat_ID & " ORDER BY k.lfdnr"
  If Not rsNa.EOF Then
   If Not IsNull(rsNa!KVNr) Then
    If rsNa!KVNr <> "0000000" And LenB(rsNa!KVNr) <> 0 Then
     obalt = 0
     For i = 0 To stand
      If Left$(Üw1(0, i), 7) = rsNa!KVNr Then
       Üw1(3, i) = Üw1(3, i) & ", HA"
       obalt = True
       Exit For
      End If
     Next i
     If Not obalt Then
      stand = stand + 1
      ReDim Preserve Üw1(4, stand) ' 0 = KV-Nr, 1 = Nachname, 2 = Vorname, 3 = Position
      Üw1ini = True
      Üw1(0, stand) = rsNa!KVNr
      Üw1(3, stand) = "HA"
     End If
    End If
   End If
  End If
  
  If Not raFaEOF Or LenB(rsNa!KVNr) <> 0 Then
'   ON Error Resume Next ' Index außerhalb des Bereichs
   If Üw1ini Then
    Üw12 = Üw1(0, 0)
    If LenB(Üw12) = 0 Then Üw12 = "-"
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Üw12/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Üw12

Function testWied()
 Dim VorDat As Date, Pat_ID$
 On Error GoTo fehler
 Call GetWord
 Call AnbogVar(True)
 Pat_ID = CStr(Forms(Anmnbi)(ABPat_ID))
 VorDat = GetVorDat(Pat_ID, False)
 Call WiederHolungsBrief(Forms(Anmnbi), VorDat, Wapp)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in testWied/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' testWied()

Function GetVorDat(Pat_ID$, obStumm%, Optional obschließ%, Optional ohneÖffnen%, Optional zeitp1 As Date, Optional name$) As Date
  Const obSichtbar = -1 ' 0 ergibt das aktuelle Datum
  Dim BRz As New ADODB.Recordset
  Dim lBrNam$, WAlt As Object, US$, Spl$(), j%
  On Error GoTo fehler
  If ohneÖffnen = 0 Then If Wapp Is Nothing Then GetWord
  If ohneÖffnen <> 0 Or Not Wapp Is Nothing Then
'   Call dtbInit
'   sql = "SELECT COUNT(0) AS ct FROM `" + QMdbAkt + "`.`briefe relevant` WHERE pat_id = " & CStr(Pat_id)
   sql = "SELECT COUNT(0) AS ct FROM `briefe` WHERE (name LIKE ""%Brief an %Dr%"" OR name LIKE ""%Arztbrief%"" OR name LIKE ""Brief an HA%"" OR name LIKE ""Brief an HAe%"") AND NOT pfad LIKE '%.pdf' AND name NOT LIKE ""%Entwurf%"" AND pat_id = " & Pat_ID
   myFrag BRz, sql
'   IF Dtb.OpenRecordset(sql)!ct > 0 THEN
   
   If BRz!ct > 0 Then
'    lBrNam = Dtb.OpenRecordset("SELECT pfad FROM `" + QMdbAkt + "`.`briefe relevant` WHERE pat_id = " & CStr(Pat_id))!Pfad
    sql = "SELECT pfad, zeitpunkt, name FROM `briefe` WHERE (name LIKE ""%Brief an %Dr%"" OR name LIKE ""%Arztbrief%"" OR name LIKE ""Brief an HA%"" OR name LIKE ""Brief an HAe%"") AND NOT pfad LIKE '%.pdf' AND name NOT LIKE ""%Entwurf%"" AND pat_id = " & Pat_ID & " ORDER BY zeitpunkt DESC"
    Set BRz = Nothing
    myFrag BRz, sql
    zeitp1 = BRz!Zeitpunkt
    name = BRz!name
    If ohneÖffnen <> 0 Then
'      GetVorDat = BRz!Zeitpunkt
    Else
     lBrNam = BRz!Pfad
     lBrNam = REPLACE$(REPLACE$(lBrNam, "$\TurboMed\Dokumente", getDokPfad), "^", "")
     On Error Resume Next
'     If Wapp.Version = 0 Then Debug.Print Wapp.Version
     If Err.Number <> 0 Then GetWord
     On Error GoTo fehler
     Select Case Wapp.Version
      Case "7.0", "8.0", "9.0"
      On Error Resume Next
 ' 24.8.12: Hier COM-Schwierigkeit bei Frau Koller
      Set WAlt = Wapp.documents.Open(lBrNam, ReadOnly:=-1, Visible:=obSichtbar) ', -1, -1, -1, , , 0, , , , , obsichtbar)
      On Error GoTo fehler
      Case "10.0", "11.0", "12.0" ' ab XP
       On Error Resume Next
 ' Ausdruck.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenConflictDocument, OpenAndRepair , DocumentDirection, NoEncodingDialog)
       Set WAlt = Wapp.documents.Open(lBrNam, 0, -1, -1, , , 0, , , , -1, obSichtbar, -1, -1)
       If Err.Number > 0 Then
        Call meld("Fehler beim Öffnen des Vorbefundes: " & lBrNam, obStumm)
       End If
       On Error GoTo fehler
     End Select
     If Not WAlt Is Nothing Then
      US$ = REPLACE$(REPLACE$(REPLACE$(WAlt.Range, vbTab, Chr$(32)), vbCr, Chr$(32)), vbVerticalTab, Chr$(32))
      Spl = Split(US)
      For j = 1 To UBound(Spl)
       If IsDate(Spl(j)) Then
        GetVorDat = CDate(Spl(j))
        Exit For
       End If
      Next j
 '     Wapp.documents(Wapp.documents.Count).Close
      If obschließ Then Wapp.documents(1).Close
     End If ' not WAlt is nothing
    End If ' not isempty(walt)
   End If ' ohneöffnen<>0
  End If
  Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetVorDat/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetVorDat(pat_id&, Optional WApp) As Date

Function WiederHolungsBrief(Nachname$, VorDat As Date, Wapp, Optional dc)
 Dim i%, lBriefDat As Date
 Dim r1, r2, rz2, rn
 Dim dcName$
 
 On Error GoTo fehler
 If VorDat = CDate(0) Then GoTo schluss
 
 With Wapp
  On Error Resume Next
  dcName = dc.name
  If Err.Number <> 0 Then
   For i = 1 To .Windows.COUNT
    If InStrB(.Windows(i), Nachname) > 0 Then
     Set dc = .Windows(i)
     GoTo habDC
    End If
   Next i
   GoTo schluss
  End If
  On Error GoTo fehler

habDC:
   On Error GoTo fehler
   Set r1 = dc.Range ' WApp.Documents(dc).Content
   Set r2 = dc.Range
   Dim r3, r0
   Set r3 = dc.Range
   With r3.Find
    .clearformatting
    .Execute ("Ich berichte")
   End With
   If r3.Find.found Then
    r3.Insertafter (" erneut")
   End If
   With r1.Find
    .clearformatting
    .Execute ("Angaben auf dem Anamnese- und Untersuchungsbogen")
   End With
   If r1.Find.found Then
    Set r0 = dc.Range(r1.Start - 1, r1.Start - 1)
    r0.Insertafter (vbCr + "Zur Vorgeschichte s. Bericht vom " + Format$(VorDat, "dd/mm/yy") + ".")
   End If
   With r2.Find
    .clearformatting
    .Execute ("Verlaufsnotizen")
   End With
   If r1.Find.found And r2.Find.found Then
    Set rn = dc.Range(r1.Start, r2.Start - 1)
    rn.Font.Hidden = True
    Call dc.bookmarks.Add("Anamnese", rn)
   End If
   Dim pakt ' Aktueller Absatz
   Dim RestDa%
   For i = 1 To 19
    On Error Resume Next
    Set r2 = dc.bookmarks("ez" & i)
    If Err.Number = 0 Then
     On Error GoTo fehler
     RestDa = 0
     Set pakt = r2.Range.paragraphs(1).Next
     Do While pakt.Range.Text <> vbCr
      If Left$(pakt.Range, 1) <> vbTab And Not IsDate(Left$(pakt.Range, 8)) Then GoTo Fertig
      If IsDate(Left$(pakt.Range, 8)) And IsNumeric(Left$(pakt.Range, 8)) Then ' Ergänzung isnumeric 27.4.08 Jörger
       If CDate(Left$(pakt.Range, 8)) > VorDat Then
         RestDa = True
         GoTo Fertig
       End If
      End If
      Set pakt = pakt.Next
     Loop
Fertig:
     On Error Resume Next
     If RestDa Then
      dc.Range(r2.Range.paragraphs(1).Range.Next.Start, pakt.Range.Start).Font.Hidden = True ' 23.10.12: pakt.range.start statt pakt.range.start-1, um Leerzeile unter Überschrift zu vermeiden
     Else ' kein REsttext => auch keine Überschrift, 25.1.12
      Do While pakt.Range.Text = vbCr
       Set pakt = pakt.Next
      Loop
      Set pakt = pakt.Previous
      dc.Range(r2.Range.Start, pakt.Next.Range.Start).Font.Hidden = True
     End If
     On Error GoTo fehler
'     r2.Range.SELECT
    Else
     On Error GoTo fehler
    End If
   Next i
   
schluss:
  End With
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in WiederHolungsBrief/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' WiederHolungsBrief(frm AS Form, WApp, Optional dc)

Function ArBName$(sverz, Nachname$, Vorname$, Pat_ID$, infos$())
 Const ABv$ = "Arztbrief vom"
 Dim i%, j%, ident%
 On Error GoTo fehler
 Call VerzPrüf(sverz)
 ArBName = Nachname & " " & Vorname & ", PID " & Pat_ID & ", " & ABv & " " & Format$(Now, "DD/MM/YY hhmmss")
 For i = 0 To UBound(infos, 2)
  ident = 0
  For j = 0 To i - 1
   If infos(4, j) = infos(4, i) Then ident = 1
  Next
  If ident = 0 Then
   If infos(14, i) <> vNS Then
    ArBName = ArBName & IIf(i = 0, " an ", " und ") & "Dr." & IIf(Len(infos(9, i)) > 0, Left$(infos(9, i), 1) & ".", "") & infos(14, i)
   End If
  End If
 Next i
 For i = 0 To UBound(infos, 2)
  ident = 0
  For j = 0 To i - 1
   If infos(15, j) = infos(15, i) Then ident = 1
  Next
  If ident = 0 Then
   If infos(15, i) <> vNS Then
    ArBName = ArBName & IIf(i = 0, " an Mail ", " und ") & infos(15, i)
   End If
  End If
 Next i
 Dim FaxNr$
 FaxNr = ""
 For i = 0 To UBound(infos, 2)
  ident = 0
  For j = 0 To i - 1
   If infos(4, j) = infos(4, i) Then ident = 1
  Next
  If ident = 0 Then
   If infos(4, i) <> vNS Then
    FaxNr = FaxNr & IIf(i = 0, " an Fax ", " und ") & infos(4, i)
   End If
  End If
 Next i
 ArBName = Left$(ArBName, 255 - Len(sverz) - Len(FaxNr) - 4) & FaxNr & ".doc"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ArBName/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' ArBName

' nur in tubriefstandalone
Function ZwischenSpeichern(sverz, Nachname$, Vorname$, Pat_ID$, infos$()) 'Optional Fax$, Optional mail$)
 Const ABv$ = "Arztbrief vom"
 Dim ABName$, FNr&
 On Error GoTo fehler
 ABName = ArBName(sverz, Nachname, Vorname, Pat_ID, infos)
 If InStrB(Wapp.activedocument.name, ABv) > 0 Then
  Wapp.activedocument.Save
 Else
  On Error GoTo fehler
  Wapp.activedocument.SaveAs REPLACE$(sverz + ABName, "/", vNS)
  If FNr <> 0 Then
   On Error GoTo fehler
   Wapp.activedocument.SaveAs ABName
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "beim Speichern unter: " + ABName + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZwischenSpeichern/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ZwischenSpeichern

Function letzteMed(dc, Pat_ID$)
 Dim raDat As New ADODB.Recordset, sql$, ZZ&, aktZ&, runde%
 On Error GoTo fehler
 sql = "SELECT  MAX(zeitpunkt) AS zp FROM `medplan` WHERE pat_id = " + Pat_ID
' raDat.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 myFrag raDat, sql
 If IsNull(raDat!Zp) Then
  dc.bookmarks("Therapiedat").Range = vNS
  Exit Function
 End If
 lTherZP = raDat!Zp
 On Error GoTo fehler
 Dim Bemerkung
' dc.bookmarks.Add Name:="MedBemerkung", Range:=dc.Range(dc.bookmarks!Therapie.End + 1, dc.bookmarks!Therapie.End + 1)
' dc.bookmarks.Add Name:="MedBemerkung", Range:=dc.Range(dc.bookmarks!DMPText.Start, dc.bookmarks!DMPText.Start)
 dc.bookmarks("Therapiedat").Range = Format$(lTherZP, "dd.mm.yy") + ":"
 ''sql = "SELECT * FROM `medplan` WHERE pat_id = " + CStr(Pat_id) + " AND zeitpunkt = (SELECT  MAX(`zeitpunkt`) AS zp FROM `medplan` WHERE pat_id = " + CStr(Pat_id) + ") ORDER BY `zeitpunkt` DESC , `feld`, `feldnr`"
'' sql = "SELECT * FROM `" + QMdbAkt + "`.`Medikamente aktuell` WHERE `medplan`.pat_id = " + CStr(Pat_id) + " ORDER BY anal, ins, metf, glib, gluci, shglin, hypt, hmg, thro, antib DESC"
' sql = "SELECT *, `medplan`.medikament AS mmedikament FROM `medplan` LEFT JOIN `medarten` ON `medplan`.medikament = `medarten`.medikament WHERE mpnr = (SELECT MAX(mpnr) FROM `medplan` WHERE zeitpunkt = (SELECT MAX(zeitpunkt) FROM `medplan` WHERE pat_id = " + CStr(Pat_id) + ")) ORDER BY anal, ins, metf, glib, gluci, shglin, hypt, hmg, thro, antib DESC"
 'sql = "SELECT `medarten`.*, `medplan`.medikament AS mmedikament,bemerkung,mo,mi,nm,ab,zn,--bbed AS j_bbed FROM `medplan` LEFT JOIN `medarten` ON `medplan`.medanfang = `medarten`.medikament WHERE `medplan`.pat_id = " & Pat_id & " AND `medplan`.mpnr = (SELECT MAX(mpnr) FROM `medplan` WHERE pat_id = " & Pat_id & " AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM `medplan` WHERE pat_id = " & CStr(Pat_id) & ")) AND NOT ISNULL(`medplan`.medikament) AND `medplan`.medikament <> '' AND `medplan`.pat_id = " & Pat_id
 ' sql = "SELECT ma.*, m.medikament mmedikament,COALESCE(bemerkung,'') Bmkg,mo,mi,nm,ab,zn,--bbed j_bbed FROM `medplan` m LEFT JOIN `medarten` ma ON m.medanfang=ma.medikament WHERE m.mpnr=(SELECT MAX(mpnr) FROM `medplan` WHERE pat_id=m.pat_id AND zeitpunkt=(SELECT MAX(zeitpunkt) FROM `medplan` WHERE pat_id = m.pat_id)) AND m.zeitpunkt=(SELECT MAX(zeitpunkt) FROM `medplan` WHERE pat_id = m.pat_id) AND NOT ISNULL(m.medikament) AND m.pat_id = " & Pat_id
 sql = "SELECT ma.*,lmp.medikament mmedikament,COALESCE(bemerkung,'') Bmkg,mo,mi,nm,ab,zn,--bbed j_bbed FROM lmp LEFT JOIN medarten ma ON lmp.medanfang=ma.medikament WHERE lmp.pat_id = " & Pat_ID
' SET radat = Dtb.OpenRecordset(sql)
 Set raDat = Nothing
 myFrag raDat, "SELECT COUNT(0) ct FROM (" & sql & ") i"
 If Not raDat.BOF Then
   ZZ = raDat!ct
 End If
 raDat.Close
' raDat.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 myFrag raDat, sql, adOpenStatic
 
 Dim Tabl
 Set Tabl = dc.Tables.Add(Range:=dc.bookmarks!Therapie.Range, NumRows:=ZZ + 1, NumColumns:=7, _
               DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitContent)
' raDat.MoveFirst
 runde = 2
' Dim beme$
 Dim bemerk$
 Do While Not raDat.EOF
'  beme = ""
  bemerk = IIf(IsNull(raDat!Bmkg), "", raDat!Bmkg)
  If bemerk <> vNS Then
'  If IsNull(raDat!Bmkg) Or raDat!Bmkg = vNS Then
'    beme = vNS
'  Else
' 25.3.23: Klammern weggelassen
   If Left$(bemerk, 2) = Chr(13) & Chr(10) Then bemerk = Mid$(bemerk, 3)
'   dc.bookmarks!MedBemerkung.Range = raDat!Bmkg
'    dc.Range(Tabl.Range.END, dc.bookmarks!DMPText.Start).Font.size = 8
'    dc.Range(Tabl.Range.END, Tabl.Range.END) = replace$(bemerk, Chr(171), "½") + Chr(13) + Chr(10)
  End If
  If IsNull(raDat!mmedikament.Value) Then Tabl.cell(runde, 1) = vNS Else Tabl.cell(runde, 1) = raDat!mmedikament.Value + bemerk
  Tabl.cell(runde, 2).Range = IIf(IsNull(raDat!mo), vNS, REPLACE$(REPLACE$(raDat!mo, "«", "½"), "¬", "¼"))
  Tabl.cell(runde, 3).Range = IIf(IsNull(raDat!mi), vNS, REPLACE$(REPLACE$(raDat!mi, "«", "½"), "¬", "¼"))
  Tabl.cell(runde, 4).Range = IIf(IsNull(raDat!nm), vNS, REPLACE$(REPLACE$(raDat!nm, "«", "½"), "¬", "¼"))
  Tabl.cell(runde, 5).Range = IIf(IsNull(raDat!ab), vNS, REPLACE$(REPLACE$(raDat!ab, "«", "½"), "¬", "¼"))
  Tabl.cell(runde, 6).Range = IIf(IsNull(raDat!Zn), vNS, REPLACE$(REPLACE$(raDat!Zn, "«", "½"), "¬", "¼"))
  Tabl.cell(runde, 7).Range = IIf(raDat!j_bBed = 0, vNS, "X")
  runde = runde + 1
  raDat.Move 1
 Loop ' While Not raDat.EOF
 Tabl.cell(1, 1) = "Medikament"
 Tabl.cell(1, 2) = "früh"
 Tabl.cell(1, 3) = "mittags"
 Tabl.cell(1, 4) = "nachm"
 Tabl.cell(1, 5) = "abends"
 Tabl.cell(1, 6) = "spät"
 Tabl.cell(1, 7) = "b.B."
 Tabl.Rows(1).Range.Font.bold = True
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in letzteMed/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' letzteMed(dc, Pat_id&)

#If False Then
Sub MedArtenUnausgefüllt()
  Const Abfr$ = "Medarten unausgefüllt"
  Dim qd$, i%, Feld1%, rs
  On Error GoTo fehler
  qd = "SELECT * FROM `" + QMdbAkt + "`.`medarten` WHERE not ("
   For i = 0 To Dtb.TableDefs!Medarten.Fields.COUNT - 1
     Select Case Dtb.TableDefs!Medarten.Fields(i).name
      Case "falsch", "false"
'       Debug.Print "ausgespart: " + Dtb.TableDefs!Medarten.Fields(i).name
      Case Else
       If Dtb.TableDefs!Medarten.Fields(i).Type = 1 Then
        qd = qd + IIf(Feld1, " or", vNS) & vNS & "`" + Dtb.TableDefs!Medarten.Fields(i).name + "`"
        Feld1 = -1
       End If
     End Select
   Next i
 qd = qd + ") ORDER BY hinzugefügt DESC"
 Call DoCmd.Close(acQuery, Abfr, acSavePrompt)
 On Error Resume Next
' DoCmd.DeleteObject acQuery, Abfr
 DtbQuerydefsDelete Abfr
 On Error GoTo fehler
 Call DtbCreateQueryDef(Abfr, qd)
#If False Then
 Set rs = Dtb.OpenRecordset(Abfr)
 Do While Not rs.EOF
   For i = 0 To Dtb.TableDefs!Medarten.Fields.COUNT - 1
     Select Case Dtb.TableDefs!Medarten.Fields(i).name
      Case "falsch", "false"
      Case Else
       If Dtb.TableDefs!Medarten.Fields(i).Type = 1 Then
        If rs.Fields(Dtb.TableDefs!Medarten.Fields(i).name) Then
'         Debug.Print rs!Medikament, rs.Fields(Dtb.TableDefs!Medarten.Fields(i).name).name + " positiv"
         GoTo weiter
        End If
       End If
     End Select
   Next i
'   Debug.Print rs!Medikament, " keines positiv"
weiter:
  rs.MoveNext
 Loop
#End If
 DoCmd.OpenQuery Abfr
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MedArtenUnausgefüllt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' MedArtenUnausgefüllt
#End If

' aufgerufen in: Epikrise()
Function FlagInit()
 Dim j%, k%
 On Error GoTo fehler
 For j = 0 To FEZ ' 1 würde reichen
  flag(j) = 0
  DiagSi(j) = 0
  DiagNamen(j) = vNS
  For k = 0 To maxIcd
   ic(j, k) = vNS
  Next k
  un(j) = 0
 Next j
 For j = 0 To BEZ ' 1 würde reichen, zur Sicherheit
  bflag(j) = 0
  bDiagsi(j) = vNS
  bDN(j) = vNS
  For k = 0 To bmaxIcd - 1
   bic(j, k) = vNS
  Next k
  bun(j) = 0
 Next j
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FlagInit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' FlagInit

'function testth()
' Lese.ProgStart
' Dim rs AS Recordset
' SET rs = myEFrag("SELECT COALESCE(therart,'') FROM therarten WHERE pat_id=12 LIMIT 1")
' IF rs.BOF THEN testth = "" ELSE testth = rs.Fields(0)
'End Function

Sub Epikrise(dc, Pat_ID$, VorDat As Date, lddat As Date, obStumm%)
  Dim i%
  On Error GoTo fehler
  nzw = vbCr
  Dim Epi$, Titel$, DT$, Folge$, FZ%, Begl$, bglz%, Akkusat$, Akklang$, Nominat$
  Dim rnam As New ADODB.Recordset, rsAnam As New ADODB.Recordset, rDT As New ADODB.Recordset, rDT0 As New ADODB.Recordset
  Dim obkomma%
  Dim lKrea!, GFR!, Alter%, DialZt%, DialAlter%, GebDat As Date
  Dim RRsyst%, RRdiast%, dmtyp$, Ther1$, obMBlAusgeh%, Gewicht#
  Dim rALz As New ADODB.Recordset
  
  'sql = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" AS NB FROM (" & laborAbfr & " WHERE pat_id = " & Pat_id & ") AS labor UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar, Normbereich AS NB " + _
   "FROM `laborxus` LEFT JOIN laborxwert ON laborxus.RefNr=laborxwert.RefNr " + _
   "WHERE pat_id = " + CStr(Pat_id) + " AND NOT EXISTS (SELECT * FROM `laborneu` WHERE pat_id = " & CStr(Pat_id) & " AND abkü = laborxwert.Abkü AND wert = laborxwert.wert AND zeitpunkt > laborxus.Eingang -3 AND zeitpunkt < laborxus.Eingang+6)"
  sql = "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_ID & " UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_ID & ") i GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb ORDER BY zeitpunkt DESC"
   
  On Error GoTo fehler
  Call FlagInit
'  SET rNam = TabÖff("Namen", "Pat_id")
  myFrag rnam, "SELECT * FROM `namen` WHERE pat_id = " & Pat_ID
'  SET rsAnam = TabÖff("Anamnesebogen", "Pat_id")
'  rsAnam.Seek "=", Pat_id
'  rsAnam.Open "SELECT -obmblausgeh AS j_obmblausgeh, a.* FROM `anamnesebogen` a WHERE pat_id = " & Pat_id, DBCn, adOpenStatic, adLockOptimistic
'  rsAnam.Open "SELECT IF(diabetestyp IN ('-','?',' '),dmtyp(pat_id),CONVERT(diabetestyp USING UTF8)) dmtyp, a.* FROM `anamnesebogen` a WHERE pat_id = " & Pat_id, DBCn, adOpenStatic, adLockOptimistic
  If rnam.BOF Then Exit Sub
  myFrag rsAnam, "SELECT IF(diabetestyp IN ('-','?',' '),dmtyp(pat_id),CONVERT(diabetestyp USING UTF8)) dmtyp, COALESCE(a.Ther1,'') Thereins, a.* FROM `anamnesebogen` a WHERE pat_id = " & Pat_ID, adOpenStatic, DBCn, adLockOptimistic
'  SET rDT = TabÖff("Diagnosen", "DiagSuch")
'  rNam.Seek "=", Pat_ID
  dmtyp = rsAnam!dmtyp
  Ther1 = rsAnam!Thereins
  obMBlAusgeh = rsAnam!obMBlAusgeh
  Gewicht = rsAnam!Gewicht
  GebDat = rsAnam!GebDat
  Titel = IIf(IsNull(rnam!Titel), vNS, rnam!Titel + " ")
  Akkusat = IIf(rnam!geschlecht = "w", "Frau ", "Herrn ") + Titel + rnam!Nachname
  Akklang = IIf(rnam!geschlecht = "w", "Frau ", "Herrn ") + Titel + rnam!Vorname & " " & rnam!Nachname
  Nominat = IIf(rnam!geschlecht = "w", "Frau ", "Herr ") + Titel + rnam!Nachname
'  Alter = (Now - CDate(rnam!GebDat)) * 2.73792574745373E-03 ' 1/365,24
  Alter = AlterBei(Now, CDate(rnam!GebDat))
  Epi = "Bei " + Akkusat
  Select Case dmtyp
   Case "1": DT = " Typ-1-Diabetes"
   Case "2": DT = " Typ-2-Diabetes"
   Case "s": DT = " sekundärer Diabetes"
   Case "g": DT = " Gestationsdiabetes"
   Case "p": DT = "e pathologische Glucosetoleranz"
   Case "-": DT = " "
   Case "?": DT = " Diabetes mellitus" ' kommt dann nicht mehr vor
   Case Else: DT = " Diabetes mellitus Typ " + IIf(IsNull(rsAnam!Diabetestyp), "", rsAnam!Diabetestyp) ' kommt dann nicht mehr vor
  End Select
  Select Case dmtyp
   Case "g", "p"
    Epi = Epi & " ist ein" & DT & " diagnostiziert worden"
   Case "-"
    Dim GfV$
    If IsNull(rsAnam("Grund für Vorstellung")) Then
     GfV = vNS
    Else
     GfV = rsAnam("Grund für Vorstellung") '!`Grund für Vorstellung`
    End If
   Case Else
    Epi = Epi & " ist ein" & DT & " seit " & DSeit(rsAnam) & " bekannt "
  End Select
  FZ = 0
  Select Case dmtyp
   Case "1", "2", "s", "p"
    Dim j%, k%, obunter%
    DiagNamen(1) = "Retinopathie" 'H36.0
    DiagNamen(flNep) = "Nephropathie" 'N08.3
    DiagNamen(flNiI) = "Niereninsuffizienz" 'N18*, N19
    DiagNamen(4) = "koronare Herzerkrankung" 'I25.*
    DiagNamen(5) = "Z.n. Myokardinfarkt" 'I25.2, I21*, I22*
    DiagNamen(6) = "Z.n. koronarem Bypass" 'Z95.1
    DiagNamen(flAVK) = "periphere Angiopathie" 'I79.2, I79.9
    DiagNamen(8) = "periphere arterielle Verschlusskrankheit" '  "I73.9", "I74", "I70", "I73"
    DiagNamen(9) = "Z.n. Bypassoperation" ' Z95.88
    DiagNamen(10) = "Z.n. Amputation" 'Z44.1 (=Amputation), Z97 (=Prothese)
    DiagNamen(11) = "Arteriosklerose der hirnzuführenden Arterien" '"I65.9", "I65", "I66", "I68"
    DiagNamen(12) = "Z.n. Apoplex" ' "I63.9Z", "I64", "I63")
    DiagNamen(13) = "Mesenterialischämie" ' K55.*
    DiagNamen(flHyp) = "Hyperkeratosen" 'L84, L85.9
    DiagNamen(flDFS) = "diabetisches Fußsyndrom" 'L89.*
    DiagNamen(flPNP) = "periphere Polyneuropathie" ' G63.2
    DiagNamen(17) = "diabetische Mononeuropathie" ' G59.0
    DiagNamen(18) = "autonome Neuropathie" ' G99.0
    DiagNamen(19) = "Hepatopathie" ' K71.6, K71.7, K76.9, K77.8, K76.0
    DiagNamen(20) = "Katarakt" 'H28.0
    DiagNamen(flLiph) = "Liphypertrophie" ' "L94.8", "L94", "E88"
    DiagNamen(22) = "Nekrobiose" ' L92.1
    DiagNamen(flChar) = "Diabetische Osteoarthropathie" ' "Charcot-Fuß" 'M14.6
    DiagNamen(24) = "Diabetische Arthropathie" ' M14.2
    DiagNamen(25) = "Diabetische Adhäsiopathie" 'M68.8
    DiagNamen(26) = "Erkrankung d.gastroint.Schleimhaut"
    DiagNamen(27) = "Manifestes Fußsyndrom"
    un(flNiI) = -1: un(5) = -1: un(6) = -1: un(8) = -1: un(9) = -1: un(10) = -1: un(12) = -1
    ic(1, 0) = "H36.0":
    ic(flNep, 0) = "N08.3": ic(flNiI, 0) = "N18": ic(flNiI, 1) = "N19"
    ic(4, 0) = "I25": ic(5, 0) = "I25.2": ic(5, 1) = "I21": ic(5, 2) = "I22"
    ic(6, 0) = "Z95.1"
    ic(flAVK, 0) = "I79.9": ic(flAVK, 1) = "I79.2": ic(flAVK, 2) = "I73"
    ic(8, 0) = "I73.9": ic(8, 1) = "I74": ic(8, 2) = "I770": ic(8, 3) = "I73"
    ic(9, 0) = "K55": ic(10, 0) = "Z44.1": ic(10, 1) = "Z97": ic(11, 0) = "I65": ic(11, 1) = "I66": ic(11, 2) = "I68"
    ic(12, 0) = "I64": ic(12, 1) = "I65"
    ic(13, 0) = "K55": ic(14, 0) = "L85.9": ic(flDFS, 0) = "L89"
    ic(flPNP, 0) = "G63.2": ic(17, 0) = "G59.0": ic(18, 0) = "G99.0"
    ic(19, 0) = "K71.6": ic(19, 1) = "K71.7": ic(19, 2) = "K76.0": ic(19, 3) = "K77.8": ic(19, 4) = "K76.9"
    ic(20, 0) = "H28.0"
    ic(flLiph, 0) = "L94.8": ic(flLiph, 1) = "E88": ic(22, 0) = "L92.1"
    ic(flChar, 0) = "M14.6": ic(24, 0) = "M14.2-": ic(25, 0) = "M68.8"
    ic(flUlc, 0) = "K23": ic(flUlc, 1) = "K25": ic(flUlc, 2) = "K26": ic(flUlc, 3) = "K27": ic(flUlc, 4) = "K28": ic(flUlc, 5) = "K29"
    ic(flMFF, 0) = "L89.2": ic(flMFF, 1) = "L89.3": ic(flMFF, 2) = "L89.4"
'    rDT.index = "DiagSuch"
'    rDT.Seek "=", Pat_ID
'    IF Not rDT.NoMatch THEN
'     Do Until rDT.EOF
     Set rDT = Nothing
'     rDT.Open "SELECT ICD, DiagSicherheit FROM `diagnosen` d WHERE d.pat_id = " & Pat_id & "  AND COALESCE(d.Dggel,0)=0 ", DBCn, adOpenDynamic, adLockReadOnly
     myFrag rDT, "SELECT ICD, DiagSicherheit FROM `diagnosen` d WHERE d.pat_id = " & Pat_ID
     Do While Not rDT.EOF
'      IF rDT!Pat_id <> Pat_id THEN Exit Do
      For j = 1 To FEZ
       For k = 0 To maxIcd
        If IsNull(ic(j, k)) Or ic(j, k) = vNS Then GoTo w1:
        If (InStrB(rDT!ICD, ic(j, k)) > 0 And ic(j, k) <> vNS) And Not (InStrB("ZA", rDT!DiagSicherheit) > 0 And un(j) = 0) Then ' bei UnterDiagnosen ist Z.n. relevant
         flag(j) = -1
         If rDT!DiagSicherheit = "V" Then DiagSi(j) = True
        End If
       Next
w1:
      Next j
      rDT.MoveNext
     Loop
'    END IF
    obunter = 0
    For j = 1 To FEZ
     If Not un(j) And obunter Then
      Folge = Folge + ")"
      obunter = 0
     End If
     If flag(j) And Not (j = flNep And obkNeph) Then
      If un(j) And Not un(j - 1) And flag(j - 1) Then
       Folge = Folge + " ("
       obunter = -1
      Else
       Folge = Folge + ", "
      End If
      If Not un(j) Then obunter = 0
      Folge = Folge + IIf(DiagSi(j), "V.a. ", vNS) + REPLACE$(DiagNamen(j), "Diabet", "diabet")
      FZ = FZ + 1
     End If
    Next j
    If Folge <> vNS Then Folge = Mid$(Folge, 3)
    
    If flag(flNep) And Not flag(flNiI) And obkNeph Then
       Dim tonRunde
      If Not obStumm Then
       For tonRunde = 1 To 10
        Call Sound(WinDir + "\media\Windows XP-Standard.wav")
       Next tonRunde
      End If
'     MsgBox "Nephropathie gestrichen!"
     flag(flNep) = 0 ' Unzutreffende Nephropathie streichen
     Call meld("Bei Pat." & Pat_ID & " wurde im Brief die Nephropathie gestrichen!", obStumm)
    End If
    
    bDN(bflHyt) = "arterielle Hypertonie" ' I1*
    bDN(2) = "Übergewicht" ' E65*, E66*, E67*, E68*
    bDN(3) = "Fettstoffwechselstörung" ' E78*
    bDN(bflRau) = "Nikotinabusus" 'F17.1
    bic(bflHyt, 0) = "I1*"
    bic(2, 0) = "E65*": bic(2, 1) = "E66*": bic(2, 2) = "E67*": bic(2, 3) = "E68*"
    bic(3, 0) = "E78*"
    bic(bflRau, 0) = "F17.1"
'    rDT.index = "DiagSuch"
'    rDT.Seek "=", Pat_ID
'    SET rDT = Nothing
    Set rDT = Nothing
'    rDT.Open " SELECT Pat_id, ICD, DiagSicherheit, DiagText FROM `diagnosen` d WHERE d.pat_id = " & Pat_id & "  AND COALESCE(d.Dggel,0)=0 ", DBCn, adOpenDynamic, adLockReadOnly
    Dim DiagT$, diags$
    myFrag rDT, "SELECT Pat_id, ICD, DiagSicherheit, DiagText FROM `diagnosen` d WHERE d.pat_id = " & Pat_ID
    If Not rDT.BOF Then
     Do Until rDT.EOF
       DiagT = rDT!DiagText
       diags = rDT!DiagSicherheit
       If rDT!Pat_ID <> CLng(Pat_ID) Then Exit Do
       For j = 1 To BEZ
        For k = 0 To bmaxIcd - 1
         If IsNull(bic(j, k)) Or bic(j, k) = vNS Then GoTo w2:
         If rDT!ICD Like bic(j, k) Then
 '         IF rDt!diagsicherheit <> "Z" THEN
           bflag(j) = -1
           Select Case diags
'           Case Null, vns, "G":  bDiagsi(j) = vns
            Case "V":            bDiagsi(j) = "V.a. "
            Case "Z":            bDiagsi(j) = "Z.n. "
            Case "A":            bDiagsi(j) = "Ausschluss "
            Case Else
             bDiagsi(j) = vNS
             If InStrB(DiagT, "V.a.") > 0 Then
              bDiagsi(j) = "V.a. "
             ElseIf InStrB(DiagT, "Z.n.") > 0 Then
              bDiagsi(j) = "Z.n. "
             ElseIf InStrB(DiagT, "Ausschl") > 0 Then
              bDiagsi(j) = "Ausschluss "
             End If
           End Select
'         END IF ' rDt!diagsicherheit <> "Z" THEN
         End If ' rDt!ICD LIKE bic(j, K) THEN
        Next
w2:
       Next j
       rDT.MoveNext
     Loop
    End If
    bglz = 0
    obunter = 0
    For j = 1 To BEZ
     If Not bun(j) And obunter Then
      Begl = Begl + ")"
     End If
     If bflag(j) Then
      If bun(j) And Not bun(j - 1) And bflag(j - 1) Then
       Begl = Begl + " ("
       obunter = -1
      Else
       Begl = Begl + ", "
      End If
      If Not bun(j) Then obunter = 0
      Begl = Begl + bDiagsi(j) + bDN(j)
      bglz = bglz + 1
     End If
    Next j
    If Begl <> vNS Then Begl = Mid$(Begl, 3)
    '         Folge = Folge + IIF(rDt!diagsicherheit = "V", "V.a.", vns) + DiagNamen(j)
    If FZ <> 0 Then
     Epi = Epi + IIf(FZ = 1, "mit der Folgeerkrankung ", "mit den Folgeerkrankungen ") + Folge
    Else
     Select Case dmtyp
      Case "g", "p"
      Case Else
       Epi = Epi + "ohne Hinweis auf Folgeerkrankungen"
     End Select
    End If
   Case Else ' dmtyp
  End Select
  If bglz <> 0 Then
   Epi = Epi + IIf(FZ = 0, ", mit ", " und ") + IIf(bglz = 1, "der Begleiterkrankung ", "den Begleiterkrankungen ") + Begl
  End If
  Epi = Epi + "." + nzw
' Therapieart
  Dim TherapieStr$, thrang%(1), lHbA1c!, loopHbA1c!, lHbA1cTag As Date, rAktLab As New ADODB.Recordset, Therlang$(1)
  For i = 0 To 1
   Dim rthera As New ADODB.Recordset
   Select Case i
    Case 0
     If VorDat = CDate(0) Then
      If IsNull(Ther1) Then
      Else
       TherapieStr = Ther1
      End If
     Else
      myFrag rthera, "SELECT therart FROM `therarten` tha WHERE pat_id = " & Pat_ID & " ORDER BY zp,mpnr LIMIT 1"
      If Not rthera.BOF Then
       TherapieStr = rthera.Fields(0)
      Else
       TherapieStr = vNS
      End If
'      TherapieStr = TherUmw(therart_erm(Pat_id, 0, rsAnam, VorDat))
     End If
    Case 1
'     TherapieStr = TherUmw(therart_erm(Pat_id, 0, rsAnam)) 'rsAnam!TherAkt
     myFrag rthera, "SELECT therart FROM `therarten` tha WHERE pat_id = " & Pat_ID & " ORDER BY zp DESC,mpnr DESC LIMIT 1"
     If Not rthera.BOF Then
      TherapieStr = rthera.Fields(0)
     Else
      TherapieStr = "Diät"
     End If
   End Select
   Set rthera = Nothing
   Dim thLang$, DiabTyp$
'   DiabTyp = IIf(IsNull(rsAnam!Diabetestyp), "", rsAnam!Diabetestyp)
   thrang(i) = ThaRang(TherapieStr, thLang, dmtyp) ' , DiabTyp)
   Therlang(i) = thLang
  Next i
  
  lHbA1c = 0
  lHbA1cTag = CDate(0)
#If True Then
  Dim Labs As labtyp
  alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
  Dim runde&, aktWert#
  Labs = LabPat(LA_HbA1c, CLng(Pat_ID))
  lHbA1c = MachNumerisch(Labs.WertSg)
  lHbA1cTag = Labs.Zp
#If miabrauchndenmaximalenhba1c Then
  Dim maxHbA1c!
  maxHbA1c = 0
  For runde = 0 To UBound(lab)
   If obLabI(LA_HbA1c, lab(runde)) Then
    aktWert = MachNumerisch(lab(runde).WertSg)
    If aktWert > maxHbA1c Then
     maxHbA1c = aktWert
    End If
   End If
  Next runde
#End If
#Else
''  SET rAktLab = Dtb.OpenRecordset("SELECT Wert,Zeitpunkt FROM (" + sql + ") AS sql1 WHERE abkü = ""HBA1C"" ORDER BY zeitpunkt DESC")
''  rAktLab.Open "SELECT Wert,Zeitpunkt FROM (" + sql + ") AS sql1 WHERE abkü = ""HBA1C"" ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly
'  myFrag rAktLab, "SELECT DATE(zeitpunkt) zp, wert FROM `laborneu` ln WHERE abkü RLIKE '^hba[c1]' AND CAST(wert AS decimal) < 22 AND pat_id = " & Pat_id & " UNION SELECT DATE(u.eingang) zp, w.wert FROM `laboryus` u LEFT JOIN laborywert w ON u.id = w.usid WHERE abkü RLIKE 'hba[c1]' AND CAST(wert AS decimal) < 22 AND pat_id = " & Pat_id & " GROUP BY zp ORDER BY zp DESC" ' wegen falsch eingetragener Fremdlabore gestrichen: AND einheit = '%'
'  IF Not rAktLab.BOF THEN
'   Do While Not rAktLab.EOF
'    IF NOT ISNULL(rAktLab!wert) THEN
'     Dim vorHbA1c$
'     vorHbA1c = replace$(replace$(rAktLab!wert, ".", ","), "%", vNS)
'     loopHbA1c = CDbl(IIF(IsNumeric(vorHbA1c), vorHbA1c, 0))
'     IF lHbA1c = 0 THEN
'      lHbA1c = loopHbA1c
'      lHbA1cTag = rAktLab!zp
'     END IF
'     IF loopHbA1c > maxHbA1c THEN maxHbA1c = loopHbA1c
'    END IF
'    rAktLab.Move 1
'   Loop
'  END IF
#End If
  
  Dim npLabZahl%, aktRund%
  If flag(flPNP) Then
   Set rAktLab = Nothing
'   rAktLab.Open "SELECT MAX(wert),MAX(zeitpunkt),langtext FROM (" + sql + ") AS sql1 WHERE abkü IN (""B12"",""BORG"",""BORM"",""BLOG"",""BLOM"",""IEPN_S"",""IEPN_U"",""C-ANCA"",""P-ANC_S"",""TPHA"",""HIV1/2_K"",""TSH"",""TSH-L_K"",""TSBF"",""T3"",""T4"") GROUP BY langtext", DBCn, adOpenDynamic, adLockReadOnly
   myFrag rAktLab, "SELECT MAX(wert),MAX(zeitpunkt),langtext FROM (" + sql + ") AS sql1 WHERE abkü IN (""B12"",""BORG"",""BORM"",""BLOG"",""BLOM"",""IEPN_S"",""IEPN_U"",""C-ANCA"",""P-ANC_S"",""TPHA"",""HIV1/2_K"",""TSH"",""TSH-L_K"",""TSBF"",""T3"",""T4"") GROUP BY langtext"
   If Not rAktLab.BOF Then
'    rAktLab.MoveLast
'    npLabZahl = rAktLab.RecordCount
    Set rALz = Nothing
'    Call rALz.Open("SELECT COUNT(0) AS ct FROM (" & rAktLab.source & ") AS innen", DBCn, adOpenDynamic, adLockReadOnly)
    myFrag rALz, "SELECT COUNT(0) AS ct FROM (" & rAktLab.source & ") AS innen"
    If Not rALz.BOF Then
     npLabZahl = rALz!ct
    Else
     npLabZahl = 0
    End If
    If npLabZahl > 5 Or (npLabZahl > 3 And lHbA1c < 7.5) Then
     Epi = Epi + "Zur Differentialdiagnostik der Neuropathie wurden mit untenstehenden Ergebnissen folgende LaborWerte untersucht: "
'     rAktLab.MoveFirst
     aktRund = 0
     Do While Not rAktLab.EOF
      Epi = Epi + IIf(aktRund > 0, ", ", vNS) + rAktLab!Langtext
      rAktLab.Move 1
      aktRund = aktRund + 1
     Loop
     Epi = Epi + "." + nzw
    End If
   End If
  End If
  
  Dim neLabZahl%
  If flag(flNiI) Or flag(flNep) Then
   Set rAktLab = Nothing
'   Call rAktLab.Open("SELECT MAX(wert),MAX(zeitpunkt),langtext FROM (" + sql + ") AS sql1 WHERE abkü IN (""C3"",""C4"",""IEPN_S"",""IEPN_U"",""C-ANCA"",""P-ANC_S"",""HIV1/2_K"",""GLOMQL"",""LUBARL"",""TUBAQL"",""CIC"") GROUP BY langtext", DBCn, adOpenDynamic, adLockReadOnly)
   myFrag rAktLab, "SELECT MAX(wert),MAX(zeitpunkt),langtext FROM (" + sql + ") sql1 WHERE abkü IN (""C3"",""C4"",""IEPN_S"",""IEPN_U"",""C-ANCA"",""P-ANC_S"",""HIV1/2_K"",""GLOMQL"",""LUBARL"",""TUBAQL"",""CIC"") GROUP BY langtext"
'   SET rAktLab = Dtb.OpenRecordset("SELECT MAX(wert),MAX(zeitpunkt),langtext FROM `" + QMdbAkt + "`.laborUNION WHERE pat_id = " + CStr(Pat_id) + " AND abkü IN (""C3"",""C4"",""IEPN_S"",""IEPN_U"",""C-ANCA"",""P-ANC_S"",""HIV1/2_K"",""GLOMQL"",""LUBARL"",""TUBAQL"",""CIC"") GROUP BY langtext")
   If Not rAktLab.BOF Then
 '   rAktLab.MoveLast
 '   neLabZahl = rAktLab.RecordCount
    Set rALz = Nothing
'    Call rALz.Open("SELECT COUNT(0) AS ct FROM (" & rAktLab.source & ") AS innen", DBCn, adOpenDynamic, adLockReadOnly)
    myFrag rALz, "SELECT COUNT(0) AS ct FROM (" & rAktLab.source & ") innen"
    If Not rALz.BOF Then
     neLabZahl = rALz!ct
    Else
     neLabZahl = 0
    End If
    If neLabZahl > 2 Then
     Epi = Epi + "Um andere Nephropathieursachen auszuschließen, wurden speziell bestimmt (Ergebnisse s.u.): "
'     rAktLab.MoveFirst
     aktRund = 0
     Do While Not rAktLab.EOF
      Epi = Epi + IIf(aktRund > 0, ", ", vNS) + rAktLab!Langtext
      rAktLab.Move 1
      aktRund = aktRund + 1
     Loop
     Epi = Epi + "." + nzw
    End If
   End If
  End If
' ""C3"",""C4"",
' Grenze zwischen Diagnostik und Therapie
  Epi = Epi + nzw
  If Not IsNull(Therlang(1)) And LenB(Therlang(1)) <> 0 Then
   If thrang(1) > thrang(0) Then
    Epi = Epi + Nominat + " ließ sich auf eine " + Therlang(1) + " einstellen."
   ElseIf thrang(1) < thrang(0) Then
    Epi = Epi + "Zuletzt wurde eine " + Therlang(1) + " gewählt."
   ElseIf thrang(1) = thrang(0) And thrang(1) < 5 And lHbA1c >= 7.5 Then
    Epi = Epi + "Zuletzt war " + Akkusat + " noch auf eine " + Therlang(1) + " eingestellt."
   Else
    Epi = Epi + "Die " + Therlang(1) + " wurde fortgeführt und angepasst."
   End If
   Epi = Epi + nzw
  End If
  
  If dmtyp <> "-" Then
   Epi = Epi + "Eine sinnvolle Verteilung und individuell bemessene Häufigkeit der prä- und postprandialen Blutzuckermessungen"
   If thrang(1) >= 4 Then
    Epi = Epi + " sowie die Vermeidung von Therapiefehlern bei der Insulinapplikation (Spritzorte, -technik, -zeitpunkte usw.)"
    If thrang(1) >= 5 Then Epi = Epi + " und Insulindosisanpassung an Mahlzeiten, Bewegung und Blutzuckerentgleisungen"
   End If
   Epi = Epi + " wurde angestrebt." + nzw
'  IF flag(flLiph) THEN Epi = Epi + "Durch häufiges Spritzen in dieselben Stellen waren Liphypertrophien aufgetreten, die in einigen Monaten wieder verschwinden und möglichst nicht an neuer Stelle auftreten sollten." + nzw
   If VorDat <> CDate(0) And flag(flLiph) Then Epi = Epi + "Liphypertrophien als Folge zu häufigen Spritzens in dieselben Stellen sollten wegen der dort ungleichmäßigen Insulinwirkung vermieden werden." + nzw
  End If
' Gestationsdiabetes
  If dmtyp = "g" Then
   Epi = Epi + "Für den weiteren Verlauf gelten nach Leitlinien folgende Empfehlungen: " & vbCrLf
   If Now - lddat < 60 Then
    Epi = Epi + "Monatliche Ultraschalluntersuchungen ab der 24. SSW, besondere Vorsicht mit Tokolytika etc., Entbindung in geeigneter Klinik"
    If thrang(1) >= 4 Then
     Epi = Epi + " mit Neonatologie, BZ perinatal 2-stündlich messen, Ziel 70 - 110 mg/dl"
    End If
    Epi = Epi + ", besondere Überwachung des Kindes"
    If thrang(1) >= 4 Then
     Epi = Epi + " (mit Bestimmung von Hb, Hkt, Ca, Mg, Bili, auch ohne klinische Auffälligkeit, zwischen 3.u.5. Tag), "
    End If
    Epi = Epi + " naßchemische Blutglucosebestimmung im Kreißsaal, Frühestfütterung, Stillen, Information des Kinderarztes über die positive Familienanamnese." & vbCrLf
    Epi = Epi + "Für die Patientin wird weiterhin empfohlen: " & vbCrLf
    Epi = Epi + "Nüchterne und postprandiale Blutzuckermessungen am 2. Tag nach der Entbindung, bei pathologischem Ausfall diabetologische Weiterbetreuung, ansonsten oraler Glucosetoleranztest nach 6-12 Wochen und dann spätestens alle 2 Jahre, "
    Epi = Epi + "bei erneuter Schwangerschaft oraler Glucosetoleranztest im 1. Trimenon, in der 24.-28. Woche sowie in der 32.-34. Woche, jeweils falls zuvor negativ."
   Else
    Epi = Epi + "Oraler Glucosetoleranztest spätestens alle 2 Jahre, "
    Epi = Epi + "bei erneuter Schwangerschaft oraler Glucosetoleranztest im 1. Trimenon, in der 24.-28. Woche sowie in der 32.-34. Woche, jeweils falls zuvor negativ."
   End If
  End If ' dmtyp = "g" THEN
' Schulungen
  Dim zp1$, zpl$, schulz%
  schulz = SchulzBest(Pat_ID, zp1, zpl, VorDat)
  If schulz > 1 Then
   Epi = Epi + Nominat + " nahm bei uns zwischen " + zp1 + " und " + zpl + " an " + CStr(schulz) + " Schulungen teil." + nzw
  End If
  Dim rlar As New ADODB.Recordset, Vero$, VeroZ%
  sql1 = "SELECT * FROM `eintraege` WHERE pat_id = " & Pat_ID & " AND art = ""lar"" AND NOT inhalt LIKE ""%umpe%"" AND NOT inhalt LIKE ""%nsul%"" "
  If VorDat <> CDate(0) Then sql1 = sql1 + "and Zeitpunkt >= " & DatFor_k(VorDat)
  sql1 = sql1 + " ORDER BY zeitpunkt"
  sql1 = "SELECT DISTINCT Inhalt FROM (" + sql1 + ") innen"
'  SET rLar = Dtb.OpenRecordset(sql1)
'  rlar.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
  myFrag rlar, sql1
  If Not rlar.BOF Then Epi = Epi + "Ich verordnete: "
  VeroZ = 0
  Do While Not rlar.EOF And Not rlar.BOF
   Dim VeroI$
   VeroI = rlar!Inhalt
   VeroI = REPLACE$(VeroI, "diabetes Therapie", "Diabetes-Therapie")
   VeroI = REPLACE$(VeroI, "diabetes adapt", "diabetesadapt")
   VeroI = REPLACE$(VeroI, "diabetes adapie", "diabetesadaptie")
   VeroI = REPLACE$(VeroI, "undAbsatzrolle", "und Absatzrolle")
   If InStrB(VeroI, "D:") > 0 Then
    Vero = Left$(VeroI, InStr(rlar!Inhalt, "D:") - 1)
   Else
    Vero = VeroI
   End If
   If Vero <> vNS Then
     Epi = Epi + IIf(VeroZ > 0, ", ", vNS) + Vero
     VeroZ = VeroZ + 1
   End If
   rlar.Move 1
  Loop
  If Not rlar.BOF Then Epi = Epi + "." + nzw
  
' Antibiose
  Dim obHWI%, AntibText$
  AntibText = vNS
  Set rDT = Nothing
'  rDT.Open "SELECT 0 FROM `diagnosen` d WHERE d.pat_id = " & Pat_id & " AND (d.icd LIKE ""N12%"" OR d.icd LIKE ""N30%"" OR d.icd LIKE ""N39%"") AND d.diagsicherheit <> 'A' AND COALESCE(d.Dggel,0)=0 ", DBCn, adOpenDynamic, adLockReadOnly
  myFrag rDT, "SELECT 0 FROM `diagnosen` d WHERE d.pat_id = " & Pat_ID & " AND (d.icd LIKE ""N12%"" OR d.icd LIKE ""N30%"" OR d.icd LIKE ""N39%"") AND d.diagsicherheit <> 'A'"
  If rDT.BOF Then obHWI = 0 Else obHWI = -1
  Dim rAntib As New ADODB.Recordset
  If obHWI Or flag(flMFF) <> 0 Then
   sql1 = "SELECT DISTINCT `medplan`.medikament, MIN(zeitpunkt) AS zeitpunkt FROM `medplan` LEFT JOIN `medarten` ON `medplan`.medanfang = `medarten`.medikament WHERE `medplan`.pat_id = " & Pat_ID & " AND antib<>0"
'  sql1 = "SELECT * FROM medikamente mit arten` WHERE antib AND `medplan`.pat_id = " & CStr(Pat_ID) & " "
   If VorDat <> CDate(0) Then sql1 = sql1 + " AND Zeitpunkt >= " & DatFor_k(VorDat)
   sql1 = sql1 & " GROUP BY `medplan`.medikament;"
'   SET rAntib = Dtb.OpenRecordset(sql1)
'   Call rAntib.Open(sql1, DBCn, adOpenDynamic, adLockReadOnly)
   myFrag rAntib, sql1
   obkomma = 0
   Do While Not rAntib.EOF
    If obkomma Then
     AntibText = AntibText + ", "
    Else
     If obHWI And Not flag(flMFF) Then
      AntibText = AntibText + "Bei der Diagnostik bezüglich diabetischer Nephropathie fiel ein Harnwegsinfekt auf, ich gab "
'     ElseIf (Not obHWI AND flag(flMFF)) THEN
'      AntibText = AntibText + "Zur Behandlung des Fußsyndroms gab ich "
     Else
      AntibText = AntibText + "Als Antibiose gab ich "
     End If
    End If
    On Error Resume Next
    Err.Clear
'    IF lies.obMySQL THEN
     AntibText = AntibText + rAntib!Medikament + " ab " + Format$(rAntib!Zeitpunkt, "dd/mm/yy")
     If Err.Number <> 0 Then
'    Else
     AntibText = AntibText + rAntib![medplan.medikament] + " ab " + Format$(rAntib!Zeitpunkt, "dd/mm/yy")
    End If
    On Error GoTo fehler
    obkomma = -1
    rAntib.Move 1
   Loop
   If obkomma Then AntibText = AntibText + "." + nzw
  End If
  
  If flag(flMFF) Then
   Epi = Epi & "Das offene diabetische Fußsyndrom wurde nach den üblichen Regeln behandelt (Druckentlastung, Überprüfung der Durchblutung" & IIf(flag(flMFF) And LenB(AntibText) <> 0, ", Antibiose", vNS) & ")." & vbCrLf
   If AntibText <> vNS Then
    Epi = Epi & AntibText & vbCrLf
   End If
  End If
  
  Dim obEintr%, mbl As New ADODB.Recordset
  Set mbl = Nothing
'  mbl.Open "SELECT art FROM `eintraege` WHERE pat_id = " & Pat_id & " AND inhalt LIKE ""%merkblatt%"" AND (inhalt LIKE ""%dfs%"" OR inhalt LIKE ""%fuß%"")", DBCn, adOpenDynamic, adLockReadOnly
  myFrag mbl, "SELECT art FROM `eintraege` WHERE pat_id = " & Pat_ID & " AND inhalt LIKE ""%merkblatt%"" AND (inhalt LIKE ""%dfs%"" OR inhalt LIKE ""%fuß%"")"
'  IF rsAnam!j_obMBlAusgeh <> 0 OR Not mbl.BOF THEN
  If obMBlAusgeh <> 0 Or Not mbl.BOF Then
   If flag(flDFS) Then
     Epi = Epi + "Wegen des diabetischen Fußsyndroms wurde " + Akkusat + " in der Vorbeugung weiterer Läsionen unterwiesen."
   Else
    Epi = Epi + "Wegen "
    If flag(flAVK) Then Epi = Epi + "der peripheren Angiopathie "
    If flag(flPNP) Then
     If flag(flAVK) Then Epi = Epi + "und "
     Epi = Epi + "der peripheren Polyneuropathie "
     If flag(flChar) Then Epi = Epi + "und der vorliegenden Diabetischen Osteoarthropathie "
     If flag(flHyp) Then Epi = Epi + "sowie Hyperkeratosen "
     If flag(flAVK) Then
      Epi = Epi + "besteht eine erhöhte Gefährdung für das diabetische Fußsyndrom. Wir haben "
     Else
      Epi = Epi + "wurde "
     End If
    End If
    If Not flag(flAVK) And Not flag(flPNP) Then
     Epi = Epi + "der Gefahr eines diabetischen Fußsyndroms wurde "
    End If
    Epi = Epi + Akkusat + " in der Vorbeugung von " + IIf(flAVK Xor flPNP, "Fußl", "L") + "äsionen unterwiesen."
   End If ' flag(flDFS) Then else
   Epi = Epi + nzw
  End If ' rsAnam!obMBlAusgeh <> 0 Or Not mbl.BOF Then
  Call GetPrRR(Pat_ID, rsAnam, RRsyst, RRdiast)
  If diI("I10", Pat_ID) Then
   Epi = Epi + nzw + "Blutdruckziel nach Leitlinien und ("
'   Epi = Epi + "Für den Blutdruck gilt nach den aktuellen Leitlinien "
   If flag(flNep) Or flag(flNiI) Then
    Epi = Epi + "für bestehende Nephropathie "
    If flag(flNiI) Then Epi = Epi + "bzw. Niereninsuffizienz "
   Else
    Epi = Epi + "für nicht nachweisbare Nephropathie"
   End If
'   Epi = Epi + " ein Zielwert von "
   Epi = Epi + "): "
' Kommentar 25.1.12
'   IF flag(flNep) OR flag(flNiI) THEN
'    Epi = Epi + "120/80 mm Hg, falls verträglich, andernfalls "
'   END IF
   Epi = Epi + " < 140/90 mm Hg (morgendl. Praxismessung)."
   If rrEmpf(RRsyst, RRdiast, Pat_ID) = "halten" Then
    Epi = Epi + " Dieses schien zuletzt erreicht worden zu sein."
   End If
   Epi = Epi + nzw
  End If
  If obMakroAlb And Alter < 61 Then
   Epi = Epi + "Dieses ist angesichts der bereits vorliegenden Makroalbuminurie und des noch relativ geringen Lebensalters besonders wichtig." + nzw
  End If
  
  Dim lLDL!, lLDLTag As Date
  lLDL = 0
  lLDLTag = CDate(0)
'  SET rAktLab = Nothing
'  Call rAktLab.Open("SELECT Wert,Zeitpunkt FROM (" + sql + ") AS sql1 WHERE abkü IN ('LDL','LDLB','LDLH','LDLC','LDLH01','LDLLG','LDLS') ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly)
  alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
  Labs = LabPat(LA_LDL, CLng(Pat_ID))
  lLDL = MachNumerisch(Labs.WertSg)
  lLDLTag = Labs.Zp
'  SET rAktLab = LabEPat(LDL, Pat_id)
''  SET rAktLab = Dtb.OpenRecordset("SELECT Wert,Zeitpunkt FROM `" + QMdbAkt + "`.LaborUNION WHERE pat_id = " + CStr(Pat_id) + " AND (abkü = ""LDL"" OR abkü = ""LDLH"") ORDER BY zeitpunkt DESC")
'  IF Not rAktLab.EOF THEN
'   Dim rAktLabZWI$
'   rAktLabZWI = Trim$(replace$(rAktLab!wert, ".", ","))
'   IF rAktLabZWI = vNS THEN
'    lLDL = 0
'   Else
'    lLDL = 0
'    ON Error Resume Next
'    lLDL = CDbl(rAktLabZWI)
'    ON Error GoTo fehler
'   END IF
'   lLDLTag = rAktLab!zp
'  END IF
' LDL
  If lLDL > 100 And Alter < 76 Then
'   Epi = Epi + "Für das LDL gilt nach den noch nicht aktualisierten Leitlinien ein Zielwert von 100 mg/dl." + nzw
   Epi = Epi + "LDL-Zielwert nach Leitlinien: 100 mg/dl (Untergrenze nicht belegbar)." + nzw
  End If
' ASS
  Dim AssA As Boolean, AssE As Boolean, AntikoagA As Boolean, AntikoagE As Boolean, keinAss
  keinAss = True ' statt 0 seit 25.1.12, nur noch bei KHK, Myokardinfarkt, Apoplex, Arteriosklerose, Nephropathie (N08.3, noch nicht aufgenommen)
'  rDT0.Open "SELECT d.icd AS icd, diagtext,g1.gruppe dg1 FROM `diagnosen` d LEFT JOIN `diagreihe` dr ON d.icd = dr.icd  LEFT JOIN diagg1 g1 ON g1.lfdnr=dr.gi1 WHERE d.pat_id = " & Pat_id & " AND (d.icd LIKE 'I25%' OR d.icd LIKE 'I21%' OR d.icd LIKE 'I20%' OR d.icd LIKE 'Z95%' OR d.icd LIKE 'I64%' OR d.icd LIKE 'I63%' OR d.icd LIKE 'I67%') AND d.diagsicherheit <> 'A' AND COALESCE(d.Dggel,0)=0 ", DBCn, adOpenDynamic, adLockReadOnly
  myFrag rDT0, "SELECT d.icd AS icd, diagtext,g1.gruppe dg1 FROM `diagnosen` d LEFT JOIN `diagreihe` dr ON d.icd = dr.icd  LEFT JOIN diagg1 g1 ON g1.lfdnr=dr.gi1 WHERE d.pat_id = " & Pat_ID & " AND (d.icd LIKE 'I25%' OR d.icd LIKE 'I21%' OR d.icd LIKE 'I20%' OR d.icd LIKE 'Z95%' OR d.icd LIKE 'I64%' OR d.icd LIKE 'I63%' OR d.icd LIKE 'I67%') AND d.diagsicherheit <> 'A'"
  If Not rDT0.BOF Then keinAss = 0
  If flag(flUlc) Then keinAss = True
  If keinAss = 0 Then
   Call TherAuskunft(Pat_ID, 1, , , , , , , , , , , , , , , , , AssA, AntikoagA)
   If AssA Then
    keinAss = True
   Else
    Call TherAuskunft(Pat_ID, 0, , , , , , , , , , , , , , , , , AssE, AntikoagE)
    If AssE Or AntikoagE Then
     keinAss = True
    Else
     Set rDT = Nothing
'     rDT.Open "SELECT d.icd AS icd, diagtext,g1.gruppe dg1 FROM `diagnosen` d LEFT JOIN `diagreihe` dr ON d.icd = dr.icd LEFT JOIN diagg1 g1 ON g1.lfdnr=dr.gi1 WHERE pat_id = " & Pat_id & " AND d.diagsicherheit <> 'A' AND COALESCE(d.Dggel,0)=0 ", DBCn, adOpenDynamic, adLockReadOnly
     myFrag rDT, "SELECT icd, COALESCE(DiagText,'') DiagText, dg1 FROM (SELECT d.icd,DiagText,g1.gruppe dg1 FROM `diagnosen` d LEFT JOIN `diagreihe` dr ON d.icd = dr.icd LEFT JOIN diagg1 g1 ON g1.lfdnr=dr.gi1 WHERE pat_id = " & Pat_ID & " AND d.diagsicherheit <> 'A') i", adOpenStatic
     Do While Not rDT.EOF
  ' Leberkrankheiten außer Leberzirrhose, Obstipation
      If rDT!ICD Like "C*" Or (rDT!dg1 = "Magen-Darm-Leber" And Not (rDT!ICD Like "K7*" And Not rDT!ICD Like "K74*") And Not rDT!ICD Like "K59*") Or InStrB(rDT!DiagText, "lutu") > 0 Or InStrB(rDT!DiagText, "sthma") > 0 Then
       keinAss = True
       Exit Do
      End If
      rDT.Move 1
     Loop
     If keinAss = 0 Then
'      IF dmtyp <> "-" AND (Now - rsAnam!GebDat) > 365 * 40 THEN
'       Epi = Epi + Nominat + " gehört zum Personenkreis derer, die (aufgrund der Diagnose " & rDT0!DiagText & ") rein statistisch gesehen von einer Thrombozytenhemmung profitieren könnten." + nzw
'      END IF
     End If ' keinAss = 0 Then
    End If ' AssE Or AntikoagE Then
   End If ' AssA Then else
  End If ' keinAss = 0 Then
  
  If flag(flNiI) Then
'   SET rAktLab = Nothing
''   Call rAktLab.Open("SELECT Wert,Zeitpunkt FROM (" + sql + ") AS sql1 WHERE abkü IN (""CREAT"",""KRE02"", ""KREA"", ""KREA02"", ""KRES"") ORDER BY zeitpunkt DESC", DBCn, adOpenDynamic, adLockReadOnly)
'   SET rAktLab = LabEPat(Krea, Pat_id)
''   SET rAktLab = Dtb.OpenRecordset("SELECT Wert,Zeitpunkt FROM `" + QMdbAkt + "`.laborUNION WHERE pat_id = " + CStr(Pat_id) + " AND abkü = ""KREA"" ORDER BY zeitpunkt DESC")
'   IF rAktLab.EOF THEN GoTo keineGFR
   alt_la = LabArt0 ' falls schon mal mit gleichen Parametern aufgerufen
   Labs = LabPat(LA_Krea, CLng(Pat_ID))
   If Labs.Zp = -1 Or Labs.Zp = #12/29/1899# Then GoTo keineGFR
   Dim sKrea$
   sKrea = MachNumerisch(Labs.WertSg)

   Dim obSchonDial%, obSchonNephr%
   Dim rDial As New ADODB.Recordset
   Dim obNiereWichtig%, obNephrologe%
     obSchonDial = 0
     obSchonNephr = 0
'   sKrea = replace$(rAktLab!wert, ".", ",")
'   IF sKrea = vNS THEN sKrea = "0"
   lKrea = CDbl(sKrea)
   If lKrea <> 0 Then
    GFR = (140 - Alter) * Gewicht * IIf(Gewicht < 3, 100, 1) * IIf(rnam!geschlecht = "w", 0.85, 1) / lKrea / 72
    DialZt = (GFR - 10) / 3
    DialAlter = Alter + DialZt
    obNiereWichtig = (DialAlter < 80)
    obNephrologe = (GFR < 50)
    If obNiereWichtig Or obNephrologe Then
     Set rDial = Nothing
     myFrag rDial, "SELECT inhalt FROM `eintraege` WHERE pat_id = " & Pat_ID & " AND instr(inhalt,'dialyse')>0"
     If Not rDial.BOF Then
      obSchonDial = True
      obNephrologe = False
     Else
      Set rDial = Nothing
      myFrag rDial, "SELECT 0 FROM `diagnosen` d WHERE d.pat_id = " & Pat_ID & " AND d.icd LIKE 'Z49%' AND d.diagsicherheit<>'A'"
      If Not rDial.BOF Then obSchonDial = True: obNephrologe = False
     End If
    End If
    If obNephrologe Then
     Set rDial = Nothing
     myFrag rDial, "SELECT 0 FROM briefe WHERE pat_id = " & Pat_ID & " AND (name LIKE '%Arenz%' OR name LIKE '%Stenglein%' OR name LIKE '%Habersetzer%' OR name LIKE '%Nephrol%')"
     If Not rDial.BOF Then obNephrologe = False
    End If
   
    If Not obSchonDial Then
     If obNiereWichtig Then
      Epi = Epi + "Bei einer geschätzten GFR von ca. " + CStr(Int(GFR)) + " ml/min "
'    IF DialAlter < 60 THEN
'     Epi = Epi + "ist mit dem Erleben von Nierenversagen zu rechnen."
'    Else
'     Epi = Epi + "könnte mit dem Erleben von Nierenversagen zu rechnen sein."
'    END IF
      Epi = Epi + "gilt es, deren Abnahmegeschwindigkeit durch Optimalhalten " + IIf(bflag(bflHyt), "von Blutdruck und Blutzucker", "des Blutzuckers") + IIf(bflag(bflRau) And LenB(bDiagsi(bflRau)) = 0, " und Einstellen des Rauchens", vNS) + " zu minimieren."
      Epi = Epi + IIf(dmtyp = "1", " Bei", " Nur bei") + " Typ-1-Diabetikern ist außerdem der Nutzen einer Eiweißrestriktion belegt."
     Else
      Epi = Epi + "Die GFR-Schätzung nach Cockcroft-Gault beträgt (mit einem Kreatinin von " + CStr(lKrea) + " mg/dl) ca. " + CStr(Int(GFR)) + " ml/min."
     End If
    End If

   Epi = Epi + nzw
   If obNephrologe Then
    If GFR < 30 Then
     Epi = Epi + "Eine weitere nephrologische Mitbehandlung halte ich für sinnvoll."
    Else
     Epi = Epi + "Eine nephrologische Mitbehandlung könnte sinnvoll sein."
    End If
   End If
   Epi = Epi + nzw
   End If ' Krea > 0
keineGFR:
  End If ' flag(flNiI) THEN
  
  Dim rUebw As New ADODB.Recordset
  sql1 = "SELECT CASE when COUNT(0)>0 THEN ue ELSE '' END FROM (" & vbCrLf & _
         "SELECT CONCAT(CHR(13),CHR(10),'Weitere Überweisungen wurden von mir ausgestellt: ',GROUP_CONCAT(uetxt ORDER BY zeitpunkt SEPARATOR ', ')) ue FROM (" & vbCrLf & _
         "SELECT pat_id,zeitpunkt," & vbCrLf & _
         "CONCAT(GROUP_CONCAT(case feldnr WHEN 0 THEN CASE feld WHEN 'Ueberweisung_an' THEN CONCAT(feldinh,' (',DATE_FORMAT(zeitpunkt,'%e.%c.%y')) when 'Diagnose' THEN CONCAT(': ',feldinh) ELSE CONCAT('; ',feldinh) END ELSE CONCAT(' ',feldinh) END ORDER BY feld DESC,feldnr SEPARATOR ''),')') Uetxt " & vbCrLf & _
         "FROM formular " & vbCrLf & _
         "WHERE feld IN ('Ueberweisung_an', 'Diagnose','DfAuftrag') " & vbCrLf & _
         "and form_abk='uew' AND formvorl RLIKE 'Überweisung gezielt' " & vbCrLf & _
         "and pat_id=" & Pat_ID & " AND zeitpunkt>" & DatFor_k(VorDat) & vbCrLf & _
         "GROUP BY pat_id,form_id,zeitpunkt " & vbCrLf & _
         ") i GROUP BY pat_id) i"
'  myEFrag "SET GROUP_CONCAT_MAX_LEN = 1000000;"
'  Epi = Epi + myEFrag(sql1).Fields(0)
  Dim rs As ADODB.Recordset
  Epi = Epi + myFrag(rs, sql1, adOpenUnspecified, Nothing, adLockReadOnly, 1000000).Fields(0)
  sql1 = "SELECT IF(COUNT(0)=0,'',CONCAT(CHR(13),CHR(10),CHR(13),CHR(10),'Weitere Termine haben wir noch vereinbart, Ihr freundliches Einverständnis voraussetzend: '," & vbCrLf & _
         "GROUP_CONCAT(CONCAT(DATE_FORMAT(zp,'%d.%m.%y %H:%i'),IF(raum RLIKE 'abor|^Seminar',CONCAT(' (',raum,')'),'')) SEPARATOR ', '),'.'))" & vbCrLf & _
         "FROM termine WHERE DATE(zp)>DATE(NOW()) AND Pid=" & Pat_ID
  Epi = Epi + myFrag(rs, sql1, adOpenUnspecified, Nothing, adLockReadOnly, 1000000).Fields(0)
  
#If False Then
  sql1 = "SELECT feldinh,zeitpunkt FROM (" & forminhalt & " WHERE pat_id = " & Pat_ID & ") AS forminhalt WHERE form_abk = ""uew"" AND feld = ""Ueberweisung_an"" "
  If VorDat <> CDate(0) Then sql1 = sql1 + "and Zeitpunkt >= " & DatFor_k(VorDat)
  sql1 = sql1 + " ORDER BY zeitpunkt"
'  rUebw.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
  myFrag rUebw, sql1
  
  If Not rUebw.BOF Then
   obkomma = 0
   Do While Not rUebw.EOF
    If Not obkomma Then Epi = Epi & vbCrLf & "Weitere Überweisungen wurden von mir ausgestellt: "
    If obkomma Then Epi = Epi + ", "
    Epi = Epi + rUebw!FeldInh + " (" + Format$(rUebw!Zeitpunkt, "dd/mm/yy")
    Dim rBegrü As New ADODB.Recordset
    'Set rBegrü = Dtb.OpenRecordset("SELECT feldinh FROM FormInhaltForm_Abk INNER JOIN " + _
                 "(`forminhaltfeldinh` INNER JOIN (`forminhaltfeld` INNER JOIN FormInhaltNeu ON " + _
                 "`forminhaltfeld`.FeldVW = FormInhaltNeu.FeldVW) ON `forminhaltfeldinh`.FeldInhVW = " + _
                 "FormInhaltNeu.FeldInhVW) ON FormInhaltForm_Abk.Form_AbkVW = FormInhaltNeu.Form_AbkVW " + _
                 "WHERE Pat_id = " + Pat_id + " AND ZeitPunkt = CDate(""" + _
                 format$(rUebw!ZeitPunkt, "DD/MM/YY hh:mm") + """) AND form_abk = ""uew"" AND Feld = ""Diagnose"" " + _
                 "ORDER BY feldnr")
'    SET rBegrü = Dtb.OpenRecordset("SELECT `Pat_ID`, `FID`, `Form_ID`, `ZeitPunkt`, `Nr`, `FeldNr`, `Feld`, `FeldInh`, form_abk FROM (((`forminhfeld` LEFT JOIN `forminhkopf` ON `forminhfeld`.`foid`=`forminhkopf`.`foid`) LEFT JOIN `formulare` ON `formulare`.`formid`=`forminhkopf`.`form_id`) LEFT JOIN `forminhaltfeld` ON `forminhfeld`.`feldvw`=`forminhaltfeld`.`feldvw`) LEFT JOIN `forminhaltfeldinh` ON `forminhfeld`.`feldinhvw`=`forminhaltfeldinh`.`feldinhvw` WHERE Pat_id = " + Pat_ID + " AND ZeitPunkt = CDate(""" + format$(rUebw!Zeitpunkt, "DD/MM/YY hh:mm") + """) AND form_abk = ""uew"" AND Feld = ""Diagnose"" " + " ORDER BY feldnr")
    Set rBegrü = Nothing
'    rBegrü.Open "SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, form_abk FROM (((`forminhfeld` LEFT JOIN `forminhkopf` ON `forminhfeld`.foid=`forminhkopf`.foid) LEFT JOIN `formulare` ON `formulare`.formid=`forminhkopf`.form_id) LEFT JOIN `forminhaltfeld` ON `forminhfeld`.feldvw=`forminhaltfeld`.feldvw) LEFT JOIN `forminhaltfeldinh` ON `forminhfeld`.feldinhvw=`forminhaltfeldinh`.feldinhvw WHERE Pat_id = " + Pat_id + " AND ZeitPunkt = " & DatFor_k(rUebw!Zeitpunkt) & " AND form_abk = ""uew"" AND Feld = ""Diagnose"" " + " ORDER BY feldnr", DBCn, adOpenDynamic, adLockReadOnly
    myFrag rBegrü, "SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, form_abk FROM (((`forminhfeld` LEFT JOIN `forminhkopf` ON `forminhfeld`.foid=`forminhkopf`.foid) LEFT JOIN `formulare` ON `formulare`.formid=`forminhkopf`.form_id) LEFT JOIN `forminhaltfeld` ON `forminhfeld`.feldvw=`forminhaltfeld`.feldvw) LEFT JOIN `forminhaltfeldinh` ON `forminhfeld`.feldinhvw=`forminhaltfeldinh`.feldinhvw WHERE Pat_id = " + Pat_ID + " AND ZeitPunkt = " & DatFor_k(rUebw!Zeitpunkt) & " AND form_abk = ""uew"" AND Feld = ""Diagnose"" " + " ORDER BY feldnr"
    If Not rBegrü.BOF Then
     Epi = Epi + ", """
     Do While Not rBegrü.EOF
      Epi = Epi + IIf(IsNull(rBegrü!FeldInh), "", rBegrü!FeldInh)
      rBegrü.Move 1
      If Not rBegrü.EOF Then Epi = Epi + " "
     Loop
     Epi = Epi + """"
    End If
    Epi = Epi + ")"
    obkomma = -1
    rUebw.Move 1
   Loop
   If obkomma Then Epi = Epi + nzw
  End If ' Not rUebw.BOF THEN
#End If
  
  Dim rKh As New ADODB.Recordset
'  SET rKH = TabÖff("KHEinweis", "Auswahl")
'  rKh.Open "SELECT * FROM kheinweis WHERE pat_id = " & Pat_id, DBCn, adOpenDynamic, adLockReadOnly
  myFrag rKh, "SELECT * FROM kheinweis WHERE pat_id = " & Pat_ID
'  rKH.Seek "=", Pat_ID
  If Not rKh.BOF Then
   obkomma = 0
   Do While Not rKh.EOF
    If rKh!Pat_ID <> CLng(Pat_ID) Then Exit Do
    If rKh!Zeitpunkt > VorDat Then
     If Not obkomma Then Epi = Epi + "Folgende Einweisungen wurden von mir ausgestellt: "
     If obkomma Then Epi = Epi + ", "
     Epi = Epi + Format$(rKh!Zeitpunkt, "dd/mm/yy") & " " & IIf(InStrB(rKh!Ziel, "einweisung") = 0, rKh!Ziel + " ", "") + _
           IIf(IsNull(rKh!Diagnose), vNS, "(" + rKh!Diagnose + ")")
     obkomma = -1
    End If ' rkh!ZeitPunkt > VorDat THEN
    rKh.Move 1
   Loop
   If obkomma Then Epi = Epi + nzw
  End If ' Not rkh.NoMatch THEN
  
  Dim rAu As New ADODB.Recordset
'  SET rAu = TabÖff("AU", "Auswahl")
'  rAu.Open "SELECT * FROM au WHERE pat_id = " & Pat_id, DBCn, adOpenDynamic, adLockReadOnly
  myFrag rAu, "SELECT * FROM au WHERE pat_id = " & Pat_ID
'  rAu.Seek "=", Pat_ID
  If Not rAu.EOF Then
   obkomma = 0
   Do
    If rAu.EOF Then Exit Do
    If rAu!Pat_ID <> CLng(Pat_ID) Then Exit Do
    If rAu!Pat_ID <> CLng(Pat_ID) Then Exit Do
    If rAu!Zeitpunkt > VorDat Then
     If Not obkomma Then Epi = Epi + "Für folgende Zeiträume wurden Arbeitsunfähigkeitsbescheinigungen ausgestellt: "
     If obkomma Then Epi = Epi + ", "
     Epi = Epi + IIf(rAu!Beginn = "00000000", vNS, Format$(Left$(rAu!Beginn, 4) + Right$(rAu!Beginn, 2), "##\.##\.##")) + "-" + Format$(Left$(rAu!Ende, 4) + Right$(rAu!Ende, 2), "##\.##\.##")
     obkomma = -1
    End If
    rAu.Move 1
   Loop
   If obkomma Then Epi = Epi + nzw
  End If ' Not rAu.NoMatch THEN
  
  
  If obHWI Then
   Epi = Epi + AntibText
  End If
  
  If dmtyp <> "-" And dmtyp <> "g" Then Epi = Epi + nzw + "Kontrolluntersuchungen gemäß dem Gesundheitspass Diabetes werden empfohlen."
  Epi = Epi + nzw
  
  dc.bookmarks("Epikrise").Range = Epi
  Dim rf 'As Range
  Dim SuchText$(), STakt
  ReDim SuchText(6)
  SuchText(0) = "Eine sinnvolle Verteilung"
  SuchText(1) = "Liphypertrophien als Folge zu häufigen"
  SuchText(2) = "Ich verordnete"
  SuchText(3) = "Blutdruckziel nach Leitlinien"
  SuchText(4) = "LDL-Zielwert nach Leitlinien"
  SuchText(5) = "Weitere Überweisungen"
  SuchText(6) = "Folgende Einweisungen"
  For Each STakt In SuchText
   Set rf = dc.Range
   With rf.Find
    .clearformatting
    .Text = STakt
    .Execute
    If .found Then
     rf.paragraphs(1).Range.Font.size = 9
    End If
   End With
  Next

  dc.bookmarks!kopfgrenze.Range = REPLACE$(REPLACE$(IIf(dmtyp = "-", dc.bookmarks!kopfgrenze.Range, REPLACE$(dc.bookmarks!kopfgrenze.Range, "Diabetologischer ", vNS)), "Patname", Akklang), "Patgeb", Format$(GebDat, "dd/mm/yyyy"))
  Exit Sub
  Dim ii%, jj%
  For ii = 1 To dc.sections.COUNT
    For jj = 1 To dc.sections(ii).Headers.COUNT
     Dim kra As Range
     Set kra = dc.Range(dc.sections(ii).Headers(jj).Range.Start, dc.bookmarks("KopfGrenze").Start)
     kra = REPLACE$(kra, "Patname", Akklang)
     kra = REPLACE$(kra, "Patgeb", Format$(rnam!GebDat, "dd/mm/yyyy"))
     If dmtyp = "-" Then dc.sections(ii).Headers(jj).Range = REPLACE$(dc.sections(ii).Headers(jj).Range, "Diabetologischer ", vNS)
    Next
  Next ii
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Epikrise/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Epikrise

'wird aufgerufen in Epikrise und DMPString
Function ThaRang(ByRef therart$, ByRef Langtext, Optional DiabTyp$) As TherapieArt
   Select Case therart ' rsAnam.Fields("Ther" + IIF(i = 0, "1", "Akt"))
    Case "Diät", "Diät?": ThaRang = Diät: Langtext = "diätetische und physikalische Therapie"
    Case "OAD":           ThaRang = OAD: Langtext = "Therapie mit oralen Antidiabetika"
    Case "GLP1":          ThaRang = glp1: Langtext = "Therapie mit GLP1-Analoga"
    Case "Komb":          ThaRang = komb: Langtext = "Kombinationstherapie mit OAD und Insulin"
    Case "GLP1Ins":       ThaRang = glp1ins: Langtext = "Kombinationstherapie mit GLP1-Analoga und Insulin"
    Case "GLP1ICT":       ThaRang = glp1ict: Langtext = "Kombination GLP1-Analoga und prand./intens.Insulintherapie"
    Case "CT", "(I)CT?", "I/CT?":
                          ThaRang = ct: Langtext = "einfache Insulintherapie"
    Case "ICT":           ThaRang = ict: Langtext = IIf(DiabTyp = "2", "prandiale Insulintherapie", "intensivierte Insulintherapie")
    Case "CSII":          ThaRang = csii: Langtext = "Insulinpumpentherapie"
   End Select
End Function ' ThaRang

Function zplschul(Pat_ID&) As Date ' Zeitpunkt letzt Schulung
  Dim sqls$, lapp As New ADODB.Recordset
  On Error GoTo fehler
  sqls$ = "SELECT MAX(zeitpunkt) zp FROM `eintraege` WHERE pat_id = " & CStr(Pat_ID) & " AND art = 'schul' "
  myFrag lapp, sqls
  If Not lapp.BOF Then
   If Not IsNull(lapp!Zp) Then
    zplschul = lapp!Zp
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
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in zplschul/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' zplschul(Pat_id&) As Date ' Zeitpunkt letzt Schulung

Function SchulzBest%(Pat_ID$, Optional zp1$, Optional zpl$, Optional abDat)
  Dim sqls$, lapp As New ADODB.Recordset
  On Error GoTo fehler
  sqls$ = "SELECT * FROM `eintraege` WHERE pat_id = " & Pat_ID + " AND art = ""schul"" "
  If Not IsMissing(abDat) Then If Now > CDate("15.10.05") Then sqls = sqls + "and Zeitpunkt >= " & DatFor_k(abDat)
  sqls = sqls + " ORDER BY zeitpunkt"
'  SET lpp = Dtb.OpenRecordset(sqls, dbOpenDynaset)
'  lapp.Open sqls, DBCn, adOpenDynamic, adLockReadOnly
  myFrag lapp, sqls, adOpenDynamic
  If Not lapp.BOF Then
   lapp.MoveLast
   zpl = Format$(lapp!Zeitpunkt, "dd/mm/yy")
   lapp.MoveFirst
   zp1 = Format$(lapp!Zeitpunkt, "dd/mm/yy")
   Set lapp = Nothing
   myFrag lapp, "SELECT COUNT(0) ct FROM (" & sqls & ") innen"
   SchulzBest = lapp!ct
  End If ' Not lapp.BOF
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

' aufgerufen in doViewsErstellen (mehrfach)
Function DiesQ()
 If Verspätung = "" Then Call holFrist
 DiesQ = " BETWEEN CAST(CONCAT(YEAR(CURRENT_TIMESTAMP() - INTERVAL " & Verspätung & " DAY),'-',(QUARTER(CURRENT_TIMESTAMP() - INTERVAL " & Verspätung & " DAY) - 1) * 3 + 1,'-01') As Date) AND CONCAT(YEAR(CURRENT_TIMESTAMP() - INTERVAL " & Verspätung & " DAY) + QUARTER(CURRENT_TIMESTAMP() - INTERVAL " & Verspätung & " DAY) DIV 4,'-',((QUARTER(CURRENT_TIMESTAMP() - INTERVAL " & Verspätung & " DAY) - 1) * 3 + 4) % 12,'-01') - INTERVAL 1 SECOND "
End Function ' DiesQ

Public Sub doViewsErstellen() ' 20.3.12: alle Views vertreten
On Error GoTo fehler
'Const Frist% = 24  ' Tage
Dim frist%, FristS$
Dim vz&
Dim VN$
Dim Vsql$
syscmd 4, "Erstelle Views für " & DBCn
vz = 0
'VN(17) = "lmp"

'vn = "_lfaelle" ' letzte Fälle bis zum Quartal der Frist
''vsql = "SELECT MAX(bhfb) AS mbhfb, pat_id, MIN(schgr) AS mschgr FROM `faelle` WHERE STR_TO_DATE(CONCAT('01.'¡(LEFT(quartal,1)-1)*3+1¡'.'¡SUBSTR(quartal,2,5)),'%d.%m.%Y') < SUBDATE(NOW(),INTERVAL " & frist & " DAY) GROUP BY pat_id ORDER BY MIN(schgr)"
'vsql = "SELECT MAX(BhFB) AS mbhfb, Pat_ID AS pid FROM `faelle` f GROUP BY Pat_ID ORDER BY pat_id"
'call DtbCreateQueryDef(vn,vsql)
'
'
'
' 19.4.20: gelöscht, da vermutlich nicht gebraucht:
'DELIMITER $$
'DROP FUNCTION IF EXISTS `bed` $$
'CREATE DEFINER=`praxis`@`%` FUNCTION `bed`() RETURNS CHAR(255) CHARSET utf8mb4 COLLATE utf8mb4_german2_ci
'    NO sql
'    DETERMINISTIC
'RETURN @bed $$
'
'DELIMITER ;

'DROP FUNCTION IF EXISTS `gesnamef` $$
'CREATE DEFINER=`praxis`@`%` FUNCTION `gesnamef`() RETURNS VARCHAR(26) CHARSET utf8mb4 COLLATE utf8mb4_german2_ci
'    DETERMINISTIC
'    Comment 'gesname für Abrechnungsprogramme'
'BEGIN
'RETURN namen.name;
'END $$

'DROP FUNCTION IF EXISTS `kontaktzahl` $$
'CREATE DEFINER=`praxis`@`%` FUNCTION `kontaktzahl`( id INT(6) UNSIGNED ) RETURNS INT(11)
'    READS SQL DATA
'    DETERMINISTIC
'    COMMENT 'kontaktzahl im aktuellen Quartal'
'BEGIN
'DECLARE kz INTEGER;
'SELECT COUNT(DISTINCT DATE(zeitpunkt)) INTO kz FROM eintraege WHERE pat_id = id AND zeitpunkt BETWEEN qanf() AND qend() AND art IN (_utf8mb4'notiz',_utf8mb4'mbf',_utf8mb4'telef',_utf8mb4'ni',_utf8mb4'gstel',_utf8mb4'gs',_utf8mb4'rz',_utf8mb4'ep',_utf8mb4'bga',_utf8mb4'tk',_utf8mb4'APK',_utf8mb4'wr',_utf8mb4'jl',_utf8mb4'ga',_utf8mb4'ih',_utf8mb4'cr',_utf8mb4'tn',_utf8mb4'hg',_utf8mb4'be',_utf8mb4'tst',_utf8mb4'ke',_utf8mb4'hz',_utf8mb4'ns',_utf8mb4'mh',_utf8mb4'ag',_utf8mb4'ph',_utf8mb4'pq',_utf8mb4'er',_utf8mb4'ds',_utf8mb4'st',_utf8mb4'eb',_utf8mb4'us',_utf8mb4'sn',_utf8mb4'vb',_utf8mb4'mip',_utf8mb4'mm',_utf8mb4'rc',_utf8mb4'ik',_utf8mb4'ks',_utf8mb4'sb',_utf8mb4'cb',_utf8mb4'th',_utf8mb4'sp',_utf8mb4'ir',_utf8mb4'as',_utf8mb4'sa',_utf8mb4'sta',_utf8mb4'eg',_utf8mb4'ans',_utf8mb4'pa',_utf8mb4'fa',_utf8mb4'fam',_utf8mb4'familie',_utf8mb4'beruf',_utf8mb4'sem',_utf8mb4'seminar',_utf8mb4'kind',_utf8mb4'bz',_utf8mb4'rp',
'_utf8mb4'B-rp',_utf8mb4'mus',_utf8mb4'muster',_utf8mb4'uzu',_utf8mb4'uz',_utf8mb4'hypo',_utf8mb4'colo',_utf8mb4'coloauf',_utf8mb4'haut',_utf8mb4'kard',
'_utf8mb4'aug ',
'_utf8mb4'auge',_utf8mb4'augen',_utf8mb4'beweg',_utf8mb4'bewegung',_utf8mb4'bew',_utf8mb4'bewg',_utf8mb4'bewe',_utf8mb4'mess',_utf8mb4'ictauf',_utf8mb4'auto',_utf8mb4'insgd',_utf8mb4'vkgd',_utf8mb4'pros',_utf8mb4'prost',_utf8mb4'prostata',_utf8mb4'uro',_utf8mb4'impf',_utf8mb4'imp',_utf8mb4'gyn',_utf8mb4'caro',_utf8mb4'carotis',_utf8mb4'ap',_utf8mb4'mu',_utf8mb4'rauch',_utf8mb4'rauchen',_utf8mb4'alko',_utf8mb4'fams',_utf8mb4'schula',_utf8mb4'ass',_utf8mb4'kra',_utf8mb4'proc',_utf8mb4'au',_utf8mb4'GPD',_utf8mb4'GDP',_utf8mb4'ba',_utf8mb4'bla',_utf8mb4'aufgd',_utf8mb4'HbA1cfre',_utf8mb4'ARCHIE2',_utf8mb4'ARCHIE',_utf8mb4'PATNRALPHA',_utf8mb4'arch',_utf8mb4'LZ',_utf8mb4'FDIAB1',_utf8mb4'FDIAB2',_utf8mb4'BLANKO',_utf8mb4'NV',_utf8mb4'M-HI',_utf8mb4'REZ-K',_utf8mb4'REZ-X',_utf8mb4'EDIAB1',_utf8mb4'LAS',_utf8mb4'EDIAB2',_utf8mb4'lab',_utf8mb4'Novo Rapid',_utf8mb4'PBA',_utf8mb4'PRAXG',_utf8mb4'FT',_utf8mb4'REZ-G',_utf8mb4'N',
'_utf8mb4'KK',_utf8mb4'REZ-P',_utf8mb4'anam',_utf8mb4'EKG01',_utf8mb4'KH',_utf8mb4'HV13',_utf8mb4'M-HM',_utf8mb4'L',_utf8mb4'IE',_utf8mb4'HÄK',_utf8mb4'tv',
'_utf8mb4 're',
'_utf8mb4'gewicht',_utf8mb4'gewi',_utf8mb4'gew',_utf8mb4'rrvgl',_utf8mb4'rrvgln',_utf8mb4'bzvgl',_utf8mb4'bzvgln',_utf8mb4'fuß',_utf8mb4'füße',_utf8mb4'taille',_utf8mb4'urin',_utf8mb4'bzm',_utf8mb4'bztp',_utf8mb4'bks',_utf8mb4'anal',_utf8mb4'andm',_utf8mb4'angd',_utf8mb4'usal',_utf8mb4'usdm',_utf8mb4'usdm1',_utf8mb4'usdm2',_utf8mb4'bauch',_utf8mb4'ufrag',_utf8mb4'uabfrag',_utf8mb4'cgma',_utf8mb4'doppler',_utf8mb4'duplex',_utf8mb4'belastun',_utf8mb4'sono',_utf8mb4'sd',_utf8mb4'UKG',_utf8mb4'Größe',_utf8mb4'groe',_utf8mb4'HbA1c',_utf8mb4'hyper',_utf8mb4'keto',_utf8mb4'wv',_utf8mb4'vw',_utf8mb4'kv',_utf8mb4'komp',_utf8mb4'kompr',_utf8mb4'debr',_utf8mb4'ulc',_utf8mb4'ulcus',_utf8mb4'EKG',_utf8mb4'LZRR',_utf8mb4'Lufu',_utf8mb4'gpt',_utf8mb4'lactoset',_utf8mb4'trop',_utf8mb4'temp',_utf8mb4'oGTT',_utf8mb4'bmi',_utf8mb4'hüfte',_utf8mb4'rr',_utf8mb4'puls',_utf8mb4'GDT',_utf8mb4'bef',_utf8mb4'OAU',_utf8mb4'OA',_utf8mb4'inj',
'_utf8mb4'inf',_utf8mb4'ht',_utf8mb4'htt',_utf8mb4'ADL',_utf8mb4'TUG',_utf8mb4'who',_utf8mb4'vac',_utf8mb4'vacc',_utf8mb4'grippe');
'RETURN kz;
'END $$



' 19.4.20: gefunden ohne Verwendungsnachweis in diesem Programm:

'DROP FUNCTION IF EXISTS `gettel` $$
'CREATE DEFINER=`praxis`@`%` FUNCTION `gettel`(dname VARCHAR(1000), anfaxstr VARCHAR(100)) RETURNS VARCHAR(1000) CHARSET utf8mb4 COLLATE utf8mb4_german2_ci
'    DETERMINISTIC
'BEGIN  DECLARE pos INT;  DECLARE ch char DEFAULT '0';  DECLARE tel VARCHAR(100) DEFAULT '';  SET pos = INSTR(dname,anfaxstr);  IF pos>0 THEN   SET pos=pos+LENGTH(anfaxstr);   wlab: LOOP    SET ch = SUBSTRING(dname,pos,1);    IF ch = '_' THEN LEAVE wlab; END IF;    IF INSTR('0123456789',ch) THEN SET tel=CONCAT(tel,ch);    ELSE IF ch='+' THEN SET tel=CONCAT(tel,'00'); END IF; END IF;    SET pos=pos+1;    IF pos>LENGTH(dname) THEN LEAVE wlab; END IF;   END loop;  END IF;  RETURN tel; END $$

'DROP FUNCTION IF EXISTS `gettel3` $$
'CREATE DEFINER=`praxis`@`localhost` FUNCTION `gettel3`(dname VARCHAR(1000), anfaxstr VARCHAR(100),anhfaxstr VARCHAR(100),ancfaxstr VARCHAR(100)) RETURNS VARCHAR(1000) CHARSET utf8mb4 COLLATE utf8mb4_german2_ci
'    DETERMINISTIC
'BEGIN
' DECLARE pos INT;
' DECLARE ch char DEFAULT '0';
' DECLARE tel VARCHAR(100) DEFAULT '';
' SET pos = INSTR(dname,anfaxstr);
' IF pos > 0 THEN
'  SET pos=pos+LENGTH(anfaxstr);
' ELSE
'  IF anhfaxstr > '' THEN
'   SET pos = INSTR(dname,anhfaxstr);
'   IF pos > 0 THEN
'    SET pos=pos+LENGTH(anhfaxstr);
'   ELSE
'    IF ancfaxstr > '' THEN
'     SET pos = INSTR(dname,ancfaxstr);
'     IF pos > 0 THEN
'      SET pos=pos+LENGTH(ancfaxstr);
'     END IF;
'    END IF;
'   END IF;
'  END IF;
' END IF;
' IF pos > 0 THEN
'wlab:   Loop
'   SET ch = SUBSTRING(dname,pos,1);
'   IF ch = '_' THEN LEAVE wlab; END IF;
'   IF INSTR('0123456789',ch) THEN SET tel=CONCAT(tel,ch);
'   ELSE IF ch='+' THEN SET tel=CONCAT(tel,'00'); END IF; END IF;
'   SET pos=pos+1;
'   IF pos>LENGTH(dname) THEN LEAVE wlab; END IF;
'  END loop;
' END IF;
' RETURN tel;
'end $$

'DROP PROCEDURE IF EXISTS `lHbAp` $$
'CREATE DEFINER=`praxis`@`%` PROCEDURE `lHbAp`(in pid INT(6) UNSIGNED, out HbA DECIMAL(9,2), out zp DATE)
'    READS SQL DATA
'    DETERMINISTIC
'    COMMENT 'letzter HbA1c'
'BEGIN
'  DECLARE zp1, zp2 DATE;
'  DECLARE Hb1, Hb2 VARCHAR(10);
'  DECLARE erg DECIMAL(9,2);
'  DECLARE done INT DEFAULT FALSE;
'  DECLARE cur1 cursOR fOR SELECT zeitpunkt, wert FROM labor1a WHERE pat_id = pid AND abkü RLIKE '^hba[1c]' AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM labor1a WHERE pat_id = pid AND abkü RLIKE '^hba[1c]');
'  DECLARE cur2 cursOR fOR SELECT zeitpunkt, wert FROM labor2a WHERE pat_id = pid AND abkü RLIKE '^hba[1c]' AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM labor2a WHERE pat_id = pid AND abkü RLIKE '^hba[1c]');
'  DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = TRUE;
'  OPEN cur1; fetch cur1 INTO zp1, Hb1; CLOSE cur1;
'  OPEN cur2; fetch cur2 INTO zp2, Hb2; CLOSE cur2;
'  IF NOT ISNULL(zp2) THEN IF zp2 > zp1 OR ISNULL(zp1) THEN SET zp1=zp2; SET Hb1=Hb2; END IF; END IF;
'  IF NOT ISNULL(zp1) THEN
'   SET erg=CONVERT(REPLACE(CONCAT('0',trim(Hb1)),',','.'),DECIMAL(9,2));
'   SET HbA=erg;
'   SET zp=zp1;
'  END IF;
'End $$

'DROP PROCEDURE IF EXISTS `lvgl` $$
'CREATE DEFINER=`praxis`@`%` PROCEDURE `lvgl`()
'    READS SQL DATA
'    COMMENT 'vergleicht laboryus mit labroneu'
'BEGIN
'drop table IF EXISTS lauswahl;
'create table lauswahl (index (befart,abkü,wert)) AS (SELECT us.id,befart,abkü,wert,0 FROM laboryus us LEFT JOIN laborywert w ON w.usid=us.id WHERE us.id=2823 ORDER BY abkü,wert);
'end $$

'DROP FUNCTION IF EXISTS `regex_repl` $$
'CREATE DEFINER=`praxis`@`%` FUNCTION `regex_repl`(pattern VARCHAR(1000),replacement VARCHAR(1000),original VARCHAR(1000)) RETURNS VARCHAR(1000) CHARSET utf8mb4 COLLATE utf8mb4_german2_ci
'    DETERMINISTIC
'BEGIN
' DECLARE temp VARCHAR(1000);
' DECLARE ch VARCHAR(1);
' DECLARE i INT;
' SET i = 1;
' SET temp = '';
' IF original REGEXP pattern THEN
'loop_label:   Loop
'   IF i > CHAR_LENGTH(original) THEN
'    LEAVE loop_label;
'   END IF;
'   SET ch = SUBSTRING(original,i,1);
'   IF NOT ch REGEXP pattern THEN
'    SET temp = CONCAT(temp,ch);
'   ELSE
'    SET temp = CONCAT(temp,replacement);
'   END IF;
'   SET i=i+1;
'  END LOOP;
' ELSE
'  SET temp = original;
' END IF;
' RETURN temp;
'END $$

'DROP FUNCTION IF EXISTS `therlang` $$
'CREATE DEFINER=`praxis`@`%` FUNCTION `therlang`(thk VARCHAR(7), pid SMALLINT(6) UNSIGNED) RETURNS VARCHAR(50) CHARSET utf8
'    READS SQL DATA
'BEGIN
'DECLARE erg VARCHAR(50) CHARACTER SET utf8;
'DECLARE dtyp VARCHAR(20);
'Case thk
' when 'Diät' THEN SET erg = 'diätetische und physikalische Therapie';
' when 'OAD' THEN SET erg = 'Therapie mit oralen Antidiabetika';
' when 'GLP1' THEN SET erg = 'Therapie mit GLP1-Analoga';
' when 'Komb' THEN SET erg = 'Kombinationstherapie mit OAD und Insulin';
' when 'GLP1Ins' THEN SET erg = 'Kombinationstherapie mit GLP1-Analoga und Insulin';
' when 'CT'  THEN SET erg = 'einfache Insulintherapie';
' when 'ICT' THEN
'    SELECT diabetestyp INTO dtyp FROM anamnesebogen WHERE pat_id = pid;
'    IF dtyp='1' THEN SET erg='intensivierte Insulintherapie';
'    ELSE SET erg = 'prandiale Insulintherapie';
'    END IF;
' when 'CSII' THEN SET erg = 'Insulinpumpentherapie';
'end case;
'RETURN erg;
'end $$

'DROP FUNCTION IF EXISTS `zuuml1` $$
'CREATE DEFINER=`praxis`@`%` FUNCTION `zuuml1`(original VARCHAR(1000)) RETURNS VARCHAR(1000) CHARSET utf8mb4 COLLATE utf8mb4_german2_ci
'    DETERMINISTIC
'BEGIN
' DECLARE erg VARCHAR(1000);
' DECLARE ch VARCHAR(1);
' DECLARE i INT;
' SET erg='';
' SET i = 1;
'loop_label:   Loop
'   IF i > CHAR_LENGTH(original) THEN
'    LEAVE loop_label;
'   END IF;
'   SET ch = SUBSTRING(original,i,1);
'   Case ch
'     when 'Ä' THEN SET erg=CONCAT(erg,'Ae');
'     when 'Ö' THEN SET erg=CONCAT(erg,'Oe');
'     when 'Ü' THEN SET erg=CONCAT(erg,'Ue');
'     when 'ä' THEN SET erg=CONCAT(erg,'ae');
'     when 'ö' THEN SET erg=CONCAT(erg,'oe');
'     when 'ü' THEN SET erg=CONCAT(erg,'ue');
'     when 'ß' THEN SET erg=CONCAT(erg,'ss');
'     ELSE SET erg=CONCAT(erg,ch);
'    END case;
'    SET i=i+1;
'   END loop;
' RETURN erg;
'END $$



Do
 FristS = InputBox("Bitte Frist eingeben:", "Frist nach Ende des Quartals zum Rückgriff auf dasselbe", 21)
 If IsNumeric(FristS) Then
  frist = CDbl(FristS)
  Verspätung = FristS
  Exit Do
 Else
  Exit Sub
 End If
Loop


'VN = "__lfaelle"
'Vsql = "SELECT BhFB mbhfb, bhfe1, Pat_ID pid FROM `faelle` f ORDER BY pat_id, bhfb;"
'Call DtbCreateQueryDef(VN, Vsql)
'vz = vz + 1


VN = "_lfaelle"
If Lese.obMySQL Then
' Vsql = "SELECT mbhfb, bhfe1, pid FROM `__lfaelle` GROUP BY pid DESC;"
' Vsql = "SELECT MIN(bhfb) bhfb, MIN(bhfe1) bhfe, pat_id pid FROM faelle GROUP BY pat_id"
 Vsql = "SELECT mbhfb,bhfe,pid FROM (" & vbCrLf & _
          "SELECT bhfb mbhfb, bhfe1 bhfe, pat_id pid " & vbCrLf & _
          ", ROW_NUMBER() OVER (PARTITION BY pat_id ORDER BY bhfb) nr FROM faelle) f " & vbCrLf & _
        "WHERE nr=1"
Else
 Vsql = "SELECT bhfb mbhfb, bhfe bhfe1,pid FROM (SELECT MIN(mbhfb) AS bhfb, MIN(bhfe1) AS bhfe, pid FROM `__lfaelle` GROUP BY pid) ORDER BY pid DESC"
End If
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "lfaelle"
'vsql = "SELECT f.* FROM `faelle` AS f INNER JOIN `_lfaelle` AS i ON f.pat_id = i.pat_id AND f.bhfb = i.mbhfb;"
' Vsql = "SELECT f.pat_id, f.fid, f.schgr, f.bhfb, f.ik, f.vknr, übwvbsnr, übwvkvnr, übwvlanr, übwr FROM `_lfaelle` l LEFT JOIN `faelle` f ON f.pat_id = l.pid AND f.bhfb = l.mbhfb"
Vsql = "SELECT pat_id, fid, schgr, bhfb, ik, vknr, übwvbsnr, übwvkvnr, übwvlanr, übwr FROM (" & vbCrLf & _
         "SELECT pat_id, fid, schgr, bhfb, ik, vknr, übwvbsnr, übwvkvnr, übwvlanr, übwr " & vbCrLf & _
         ", ROW_NUMBER() OVER (PARTITION BY pat_id ORDER BY bhfb) nr FROM faelle) f " & vbCrLf & _
       "WHERE nr=1"


Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


'VN = "labor2"
'Vsql = "SELECT u.Pat_id AS Pat_ID,CAST(u.Eingang As Date) AS zeitpunkt,u.BefArt AS FertigStGrad,w.Abkü AS Abkü,w.Langname AS Langtext,w.Wert AS Wert,w.Einheit AS Einheit,w.Kommentar AS Kommentar,w.Normbereich AS NB FROM (`laborxus` u LEFT JOIN laborxwert w on((u.RefNr = w.RefNr))) WHERE (not(exists(SELECT 1 AS Not_used FROM `laborneu` WHERE ((`laborneu`.Pat_ID = u.Pat_id) AND (`laborneu`.Abkü = w.Abkü) AND (`laborneu`.Wert = w.Wert) AND (`laborneu`.ZeitPunkt > SUBDATE(u.Eingang,INTERVAL 3 DAY)) AND (`laborneu`.ZeitPunkt < ADDDATE(u.Eingang, INTERVAL 6 DAY))))))  AND wert <> '' AND NOT ISNULL(wert) ;"
'Call DtbCreateQueryDef(VN, Vsql)
'vz = vz + 1

'(not(exists(SELECT 1 AS Not_used FROM `laborneu` WHERE ((`laborneu`.Pat_ID = u.Pat_id) AND (`laborneu`.Abkü = w.Abkü) AND (`laborneu`.Wert = w.Wert) AND (`laborneu`.ZeitPunkt > SUBDATE(u.Eingang,INTERVAL 3 DAY)) AND (`laborneu`.ZeitPunkt < ADDDATE(u.Eingang, INTERVAL 6 DAY))))))
'VN = "_labor2a"
'#IF False THEN
''Vsql = "SELECT u.Pat_id Pat_ID,CAST(u.Eingang As Date) zeitpunkt,u.BefArt FertigStGrad,w.Abkü Abkü,w.Langname Langtext,w.Wert Wert,w.Einheit Einheit, CONCAT(IF(erklärung RLIKE '^:[ /\*:]*$','',IF(erklärung RLIKE '^:[ /\*]*:',CONCAT(MID(erklärung,locate(':',erklärung,2)+1),';'),IF(erklärung='.','',IF(erklärung='','',CONCAT(erklärung,';'))))),kommentar) Kommentar, nb.NB NB, nb.uNg uNg, nb.oNg oNg, s.labor Labor FROM (`laborxus` u LEFT JOIN `laborxwert` w ON u.RefNr = w.RefNr LEFT JOIN `laborxpnb` nb ON w.nbid = nb.id) LEFT JOIN `laborxsaetze` s ON u.satzid = s.satzid WHERE wert <> '' AND NOT ISNULL(wert)"
'Vsql = "SELECT u.Pat_id Pat_ID,CAST(u.Eingang As Date) zeitpunkt,u.BefArt FertigStGrad,IF(ISNULL(bez1.abkü),w.abkü,bez1.abkü) Abkü,IF(ISNULL(bez1.abkü),w.langname,bez1.langtext) Langtext,w.Wert Wert,IF(ISNULL(bez1.abkü),w.einheit,bez1.einheit) Einheit, CONCAT(IF(erklärung RLIKE '^:[ /\*:]*$','',IF(erklärung RLIKE '^:[ /\*]*:',CONCAT(MID(erklärung,locate(':',erklärung,2)+1),';'),IF(erklärung='.','',IF(erklärung='','',CONCAT(erklärung,';'))))),kommentar) Kommentar, nb.NB NB, nb.uNg uNg, nb.oNg oNg, s.labor Labor " & _
'"FROM (`laborxus` u " & _
'"LEFT JOIN `laborxwert` w ON u.RefNr = w.RefNr " & _
'"LEFT JOIN `laborxpneu` p ON w.abkü = p.abkü AND w.einheit = p.einheit " & _
'"LEFT JOIN `laborxpgl` gl ON p.id = gl.idxpbez " & _
'"LEFT JOIN `laborxpneu` bez1 ON gl.idxpneu = bez1.id " & _
'"LEFT JOIN `laborxpnb` nb ON w.nbid = nb.id) " & _
'"LEFT JOIN `laborxsaetze` s ON u.satzid = s.satzid WHERE wert <> '' AND NOT ISNULL(wert)"
'#ELSEIF False THEN
''Const KomStr$ = "CONCAT(IF(erklärung RLIKE '^:[ /\*:]*$','',IF(erklärung RLIKE '^:[ /\*]*:',CONCAT('Erkl.:',CONCAT(MID(erklärung,locate(':',erklärung,2)+1),';')),IF(erklärung='.','',IF(erklärung='','',CONCAT(CONCAT('Erkl.:',erklärung),';'))))),CONCAT(' Kom.:',kommentar))"
'Const KomStr$ = "CONCAT(IF(erklärung RLIKE '^:[ /\*:]*$','',IF(erklärung RLIKE '^:[ /\*]*:',CONCAT(MID(erklärung,locate(':',erklärung,2)+1),';'),IF(erklärung='.','',IF(erklärung='','',CONCAT(erklärung,';'))))),kommentar)"
''Vsql = "SELECT u.Pat_id Pat_ID,CAST(u.Eingang As Date) zeitpunkt,u.BefArt FertigStGrad,w.Abkü Abkü,w.Langname Langtext,w.Wert Wert,w.Einheit Einheit, " & KomStr & " Kommentar, nb.NB NB, nb.uNg uNg, nb.oNg oNg, s.labor Labor FROM (`laborxus` u LEFT JOIN `laborxwert` w ON u.RefNr = w.RefNr LEFT JOIN `laborxpnb` nb ON w.nbid = nb.id) LEFT JOIN `laborxsaetze` s ON u.satzid = s.satzid" ' WHERE wert <> '' AND NOT ISNULL(wert)"

''Vsql = "SELECT u.Pat_id Pat_ID,CAST(u.Eingang As Date) zeitpunkt,u.BefArt FertigStGrad,IF(ISNULL(n2.abkü),w.abkü,n2.abkü) Abkü, w.Abkü abk_ur,w.Langname Langtext,w.Wert Wert,IF(ISNULL(n2.abkü),w.einheit,n2.einheit) Einheit, w.Einheit Einheit_ur," & KomStr & " Kommentar, IF(ISNULL(n2.abkü),nb.NB,nb2.nb) NB, nb.NB NB_ur, IF(ISNULL(n2.abkü),nb.uNg,nb2.uNg) uNg, nb.uNg uNg_ur, IF(IF(ISNULL(n2.abkü),w.abkü,n2.abkü) = 'LDL' AND IF(ISNULL(n2.abkü),w.einheit,n2.einheit) = 'mg/dl','100',IF(ISNULL(n2.abkü),nb.oNg,nb2.oNg)) oNg, nb.oNg oNg_ur, s.labor Labor FROM (`laborxus` u LEFT JOIN `namen` na ON u.pat_id = na.pat_id LEFT JOIN `laborxwert` w ON u.RefNr = w.RefNr LEFT JOIN `laborxpnb` nb ON w.nbid = nb.id LEFT JOIN `laborxpneu` neu ON nb.pid = neu.id LEFT JOIN `laborxpgl` gl ON neu.id = gl.idxpneu LEFT JOIN `laborxpneu` n2 ON gl.idxpbez = n2.id LEFT JOIN `laborxpnb` nb2 ON n2.id = nb2.pid " & _
'"AND nb2.geschlecht IN (3,9,0,IF(na.geschlecht='m',IF(ADDDATE(na.gebdat,INTERVAL 18 YEAR)>CAST(u.eingang As Date),1,4) ,IF(ADDDATE(na.gebdat,INTERVAL 18 YEAR)>CAST(u.eingang As Date),2,5)))  ) LEFT JOIN `laborxsaetze` s ON u.satzid = s.satzid"
'Vsql = "SELECT u.Pat_id Pat_ID,CAST(u.Eingang As Date) zeitpunkt,u.BefArt FertigStGrad,IF(ISNULL(n2.abkü),w.abkü,n2.abkü) Abkü,w.Abkü abk_ur,w.Langname Langtext,w.Wert Wert,IF(ISNULL(n2.abkü),w.einheit,n2.einheit) Einheit,w.Einheit Einheit_ur," & KomStr & " Kommentar,nb.NB NB,nb.NB NB_ur,nb.uNg uNg,nb.uNg uNG_ur,IF(IF(ISNULL(n2.abkü),w.abkü,n2.abkü) = 'LDL' AND IF(ISNULL(n2.abkü),w.einheit,n2.einheit) = 'mg/dl','100',nb.oNg) oNg,nb.oNg oNg_ur,s.Labor Labor, e.pfad " & _
       '"FROM (`laborxus` u " & _
       '"LEFT JOIN `namen` na ON u.pat_id = na.pat_id " & _
       '"LEFT JOIN `laborxwert` w ON u.RefNr = w.RefNr " & _
       '"LEFT JOIN `laborxpnb` nb ON w.nbid = nb.id " & _
       '"LEFT JOIN `laborxpneu` neu ON nb.pid = neu.id " & _
       '"LEFT JOIN `laborxpgl` gl ON neu.id = gl.idxpneu " & _
       '"LEFT JOIN `laborxpneu` n2 ON gl.idxpbez = n2.id) " & _
       '"LEFT JOIN `laborxsaetze` s ON u.satzid = s.satzid " & _
       '"LEFT JOIN `laborxeingel` e ON s.datid = e.datid " & _
       '"ORDER BY u.pat_id, u.eingang, u.BefArt"
       ''  GROUP BY  u.pat_id, u.Eingang, FertigstGrad, w.abkü, w.wert, w.einheit,nb.nb
'' WHERE (ISNULL(nb2.eingang) OR nb2.eingang = (SELECT MAX(eingang) FROM `laborxpnb` WHERE pid = n2.id))
'#ELSE
' identisch in labimp/turbomed.cpp

' das Folgende muss koordiniert werden mit labbimp/labimp.cpp!
VN = "labor2a"
Vsql = _
"SELECT * FROM (" & vbCrLf & _
"SELECT Pat_id, eingang Zeitpunkt, befart FertigStGrad, w.Abkü, w.langtext Langtext" & vbCrLf & _
",TRIM(IF(w.Abkü='ALBUM' AND Wert='' AND k.Text LIKE 'nicht berechenb%','< 20',IF(TRIM(Wert) REGEXP '^[0-9]+\\,?[0-9]*$', REPLACE(Wert,',','.'),Wert))) Wert" & vbCrLf & _
",w.Einheit " & vbCrLf & _
",CONCAT(IF(ISNULL(e.text) OR e.text RLIKE '^:[ /\\*:]*$','',IF(e.text RLIKE '^:[ /\\*]*:'" & vbCrLf & _
",CONCAT(MID(e.text,locate(':',e.text,2)+1),';'),IF(e.text='.','',IF(e.text='','',CONCAT(e.text,';'))))),IF(ISNULL(k.text),'',k.text)) Kommentar " & vbCrLf & _
",n.NB, uNg" & vbCrLf & _
",IF(w.abkü='LDL' AND w.einheit='mg/dl','100',oNg) oNg" & vbCrLf & _
",l.Labor, Pfad, d.DatID " & vbCrLf & _
",p.Gruppe, p.Reihe,2 Qu " & vbCrLf & _
"FROM laboryus u " & vbCrLf & _
"LEFT JOIN laborywert w ON w.usid=u.id " & vbCrLf & _
"LEFT JOIN laboryhinw e ON e.id=w.erklid " & vbCrLf & _
"LEFT JOIN laboryhinw k ON k.id=w.kommid " & vbCrLf & _
"LEFT JOIN laborypnb n ON n.id=w.nbid " & vbCrLf & _
"LEFT JOIN laborysaetze s ON s.satzid=u.satzid " & vbCrLf & _
"LEFT JOIN laborydat d ON d.datid=s.datid " & vbCrLf & _
"LEFT JOIN laboryplab l ON l.id=s.labid " & vbCrLf & _
"LEFT JOIN laborparameter p ON p.abkü=w.abkü AND p.einheit=IF(w.einheit IN ('','\'kA\''),'kA',w.einheit)" & vbCrLf & _
"  AND p.NBm = n.NB" & vbCrLf & _
"WHERE (wert<>'' AND wert IS NOT NULL) ) i" & vbCrLf & _
"WHERE (kommentar<>'' AND kommentar IS NOT NULL)"
' "LEFT JOIN laborparameter p ON p.abkü=CAST(w.abkü AS char CHARSET utf8mb4) COLLATE utf8mb4_german2_ci AND p.einheit=CAST(IF(w.einheit IN ('','\'kA\''),'kA',w.einheit) AS char CHARSET utf8mb4) COLLATE utf8mb4_german2_ci "

'#END IF
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "labor2bakt"
Vsql = _
"SELECT Verf,KuQu,Quelle,QSpez,IF(ISNULL(k.text),'',k.text) Kommentar, IF(ISNULL(e.text),'',e.text) Erkl, IF(ISNULL(h.text),'',h.text) Hinw, Keimzahl,abrd,Auftragsnummer," & vbCrLf & _
"Auftragsschlüssel,Eingang,Berichtsdatum,Pat_id,nachname,vorname,GebDat,Zeitpunktlaborneu,Pat_id_Laborneu " & vbCrLf & _
"FROM laborybakt b " & vbCrLf & _
"LEFT JOIN laboryus u ON b.usid=u.id " & vbCrLf & _
"LEFT JOIN laboryhinw k ON b.kommid=k.id " & vbCrLf & _
"LEFT JOIN laboryhinw e ON b.erklid=e.id " & vbCrLf & _
"LEFT JOIN laboryhinw h ON b.hinwid=h.id"
'"LEFT JOIN laborysaetze s ON u.satzid=s.satzid " & vbCrLf & _
'"LEFT JOIN laborydat d ON s.datid=d.datid;"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

'Vsql = "SELECT * FROM (SELECT zeitpunkt zp, CAST(inhalt AS CHAR(100)) wert FROM `eintraege` " & vbCrLf & _
"WHERE pat_id = " & PID & " AND art LIKE 'urin%' " & vbCrLf & _
"UNION SELECT eingang zp, CAST(keimzahl AS CHAR(16)) wert FROM `laboryus` us " & vbCrLf & _
"LEFT JOIN `laborybakt` b ON us.id=b.usid WHERE pat_id = " & PID & " AND keimzahl<> '')  u2 ORDER BY zp DESC"

VN = "labor2aNachweis"
Vsql = "SELECT d.pfad, d.datid,s.satzid,u.id usid,w.id,pat_id_0,pat_id_1,pat_id_2,pat_id_3,pat_id_4,pat_id_5,pat_id_6,pat_id_7,pat_id_laborneu,Pat_id, eingang Zeitpunkt, befart FertigStGrad, w.Abkü, w.langtext Langtext" & vbCrLf & _
",trim(IF(trim(Wert) REGEXP '^[0-9]+\\,?[0-9]*$', REPLACE(Wert,',','.'),Wert)) Wert" & vbCrLf & _
",w.Einheit " & _
",CONCAT(IF(ISNULL(e.text) OR e.text RLIKE '^:[ /\\*:]*$','',IF(e.text RLIKE '^:[ /\\*]*:',CONCAT(MID(e.text,locate(':',e.text,2)+1),';'),IF(e.text='.','',IF(e.text='','',CONCAT(e.text,';'))))),IF(ISNULL(k.text),'',k.text)) Kommentar " & _
",n.NB, uNg" & vbCrLf & _
",IF(w.abkü='LDL' AND w.einheit='mg/dl','100',oNg) oNg" & vbCrLf & _
", u.Auftragsnummer, u.Auftragsschlüssel, l.Labor " & vbCrLf & _
",Gruppe, Reihe,2 Qu " & vbCrLf & _
"FROM laboryus u " & vbCrLf & _
"LEFT JOIN laborywert w ON u.id=w.usid " & vbCrLf & _
"LEFT JOIN laboryhinw e ON w.erklid=e.id " & vbCrLf & _
"LEFT JOIN laboryhinw k ON w.kommid=k.id " & vbCrLf & _
"LEFT JOIN laborypnb n ON w.nbid=n.id " & vbCrLf & _
"LEFT JOIN laborysaetze s ON u.satzid=s.satzid " & vbCrLf & _
"LEFT JOIN laborydat d ON s.datid=d.datid " & vbCrLf & _
"LEFT JOIN laboryplab l ON s.labid=l.id " & vbCrLf & _
"LEFT JOIN laborparameter p ON p.abkü=w.abkü AND p.einheit=IF(w.einheit IN ('','\'kA\''),'kA',w.einheit)" & vbCrLf & _
"  AND p.NBm = n.NB"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1
' "LEFT JOIN laborparameter p ON p.abkü=CAST(w.abkü AS char CHARSET utf8mb4) COLLATE utf8mb4_german2_ci AND p.einheit=CAST(IF(w.einheit IN ('','\'kA\''),'kA',w.einheit) AS char CHARSET utf8mb4) COLLATE utf8mb4_german2_ci "

#If False Then
'VN = "labor1"
''Vsql = "SELECT n.Pat_ID AS Pat_ID,CAST(n.ZeitPunkt As Date) AS ZeitPunkt,n.FertigStGrad AS FertigStGrad,n.Abkü AS Abkü,l.Langtext AS Langtext,n.Wert AS Wert,n.Einheit AS Einheit,k.Kommentar AS Kommentar,_utf8'' AS nb FROM (`laborneu` n LEFT JOIN `laborlangtext` l ON l.langtextvw = n.langtextvw) LEFT JOIN laborkommentar k ON k.KommentarVW = n.kommentarvw WHERE wert <> '' AND NOT ISNULL(wert) ;"
'Vsql = "SELECT n.Pat_ID Pat_ID,CAST(n.ZeitPunkt As Date) ZeitPunkt,n.FertigStGrad FertigStGrad,IF(ISNULL(n2.abkü),n.abkü,n2.abkü) Abkü, n.Abkü abk_ur,l.Langtext Langtext,IF(n.Wert REGEXP '^[0-9]+\\,?[0-9]*$', REPLACE(n.Wert,',','.'),n.Wert) Wert,IF(ISNULL(n2.abkü),n.einheit,n2.einheit) Einheit, n.Einheit Einheit_ur,k.Kommentar Kommentar, IF(ISNULL(n2.abkü),CONCAT(IF(na.geschlecht='m' OR ISNULL(uNw) OR uNw='',uNm,uNw),'-',IF(na.geschlecht='m' OR ISNULL(oNw) OR oNw = '',oNm,oNw)),nb2.nb) NB, CONCAT(IF(na.geschlecht='m' OR ISNULL(uNw) OR uNw='',uNm,uNw),'-',IF(na.geschlecht='m' OR ISNULL(oNw) OR oNw='',oNm,oNw)) NB_ur, IF(ISNULL(n2.abkü),IF(na.geschlecht='m' OR ISNULL(uNw) OR uNw='',uNm,uNw),nb2.uNg) uNg, IF(na.geschlecht='m' OR ISNULL(uNw) OR uNw='',uNm,uNw) uNg_ur, IF(ISNULL(n2.abkü),IF(na.geschlecht='m' OR ISNULL(oNw) OR oNw='',oNm,oNw),nb2.oNg) oNg, IF(na.geschlecht='m' OR ISNULL(oNw) OR oNw='',oNm,oNw) oNg_ur," & _
'"'TM' Labor, '' Pfad FROM (`laborneu` n LEFT JOIN `namen` na ON n.pat_id = na.pat_id LEFT JOIN `laborlangtext` l ON l.langtextvw = n.langtextvw) " & _
'"LEFT JOIN `laborkommentar` k ON k.KommentarVW = n.kommentarvw  LEFT JOIN `laborparameter` p ON n.abkü = p.abkü AND n.einheit = p.einheit LEFT JOIN `laborypgl` gl ON p.id = gl.idpara LEFT JOIN `laborypneu` n2 ON gl.idypbez = n2.id  LEFT JOIN `laborypnb` nb2 ON n2.id = nb2.pneuid WHERE wert <> '' AND NOT ISNULL(wert)" ' AND nb2.geschlecht IN (3,9,0,IF(na.geschlecht='m',IF(ADDDATE(na.gebdat,INTERVAL 18 YEAR)>CAST(eingang As Date),1,4) ,IF(ADDDATE(na.gebdat,INTERVAL 18 YEAR)>CAST(eingang As Date),2,5)))
'Call DtbCreateQueryDef(VN, Vsql)
'vz = vz + 1
#End If

VN = "labor1a"
'Vsql = "SELECT n.Pat_ID AS Pat_ID,CAST(n.ZeitPunkt As Date) AS ZeitPunkt,n.FertigStGrad AS FertigStGrad,n.Abkü AS Abkü,l.Langtext AS Langtext,n.Wert AS Wert,n.Einheit AS Einheit,k.Kommentar AS Kommentar,_utf8'' AS nb FROM (`laborneu` n LEFT JOIN `laborlangtext` l ON l.langtextvw = n.langtextvw) LEFT JOIN laborkommentar k ON k.KommentarVW = n.kommentarvw WHERE wert <> '' AND NOT ISNULL(wert) AND NOT EXISTS(SELECT u.Pat_id AS _Pat_ID,CAST(u.Eingang As Date) AS _zeitpunkt,u.BefArt AS _FertigStGrad,w.Abkü AS _Abkü,w.Langname AS _Langtext,w.Wert AS _Wert,w.Einheit AS _Einheit,w.Kommentar AS _Kommentar,nb.NB AS _NB FROM `laborxus` u LEFT JOIN laborxwert w ON u.RefNr = w.RefNr LEFT JOIN `laborxpnb` nb ON w.nbid = nb.id WHERE u.pat_id = n.pat_id AND w.abkü = n.abkü AND w.wert = n.wert AND CAST(n.zeitpunkt As Date) > SUBDATE(CAST(u.Eingang As Date),INTERVAL 3 DAY) AND CAST(n.zeitpunkt As Date) < ADDDATE(CAST(u.Eingang As Date),INTERVAL 6 DAY) )"
'Vsql = "SELECT n.Pat_ID AS Pat_ID,CAST(n.ZeitPunkt As Date) AS ZeitPunkt,n.FertigStGrad AS FertigStGrad,n.Abkü AS Abkü,l.Langtext AS Langtext,n.Wert AS Wert,n.Einheit AS Einheit,k.Kommentar AS Kommentar, CONCAT(IF(na.geschlecht='m' OR ISNULL(uNw),uNm,uNw),'-',IF(na.geschlecht='m' OR ISNULL(oNw),oNm,oNw)) NB, IF(na.geschlecht='m' OR ISNULL(uNw),uNm,uNw) uNg, IF(na.geschlecht='m' OR ISNULL(oNw),oNm,oNw) oNg, 'TM' Labor FROM (`laborneu` n LEFT JOIN `namen` na ON n.pat_id = na.pat_id LEFT JOIN `laborlangtext` l ON l.langtextvw = n.langtextvw) LEFT JOIN laborkommentar k ON k.KommentarVW = n.kommentarvw  LEFT JOIN laborparameter p ON n.abkü = p.abkü AND n.einheit = p.einheit WHERE wert <> '' AND NOT ISNULL(wert)" & _
 " AND NOT EXISTS(SELECT 0 FROM `laborxus` u LEFT JOIN `laborxwert` w ON u.RefNr = w.RefNr LEFT JOIN `laborxpnb` nb ON w.nbid = nb.id WHERE u.pat_id = n.pat_id AND w.abkü = n.abkü AND w.wert = n.wert AND CAST(n.zeitpunkt As Date) > SUBDATE(CAST(u.Eingang As Date),INTERVAL 3 DAY) AND CAST(n.zeitpunkt As Date) < ADDDATE(CAST(u.Eingang As Date),INTERVAL 6 DAY) )"
'Vsql = "SELECT n.Pat_ID Pat_ID,CAST(n.ZeitPunkt As Date) ZeitPunkt,n.FertigStGrad FertigStGrad,IF(ISNULL(n2.abkü),n.abkü,n2.abkü) Abkü, n.Abkü abk_ur,l.Langtext Langtext,n.Wert Wert,IF(ISNULL(n2.abkü),n.einheit,n2.einheit) Einheit, n.Einheit Einheit_ur,k.Kommentar Kommentar, IF(ISNULL(n2.abkü),CONCAT(IF(na.geschlecht='m' OR ISNULL(uNw) OR uNw='',uNm,uNw),'-',IF(na.geschlecht='m' OR ISNULL(oNw) OR oNw = '',oNm,oNw)),nb2.nb) NB, CONCAT(IF(na.geschlecht='m' OR ISNULL(uNw) OR uNw='',uNm,uNw),'-',IF(na.geschlecht='m' OR ISNULL(oNw) OR oNw='',oNm,oNw)) NB_ur, IF(ISNULL(n2.abkü),IF(na.geschlecht='m' OR ISNULL(uNw) OR uNw='',uNm,uNw),nb2.uNg) uNg, IF(na.geschlecht='m' OR ISNULL(uNw) OR uNw='',uNm,uNw) uNg_ur, IF(ISNULL(n2.abkü),IF(na.geschlecht='m' OR ISNULL(oNw) OR oNw='',oNm,oNw),nb2.oNg) oNg, IF(na.geschlecht='m' OR ISNULL(oNw) OR oNw='',oNm,oNw) oNg_ur," & _
"'TM' Labor FROM (`laborneu` n LEFT JOIN `namen` na ON n.pat_id = na.pat_id LEFT JOIN `laborlangtext` l ON l.langtextvw = n.langtextvw) " & _
"LEFT JOIN `laborkommentar` k ON k.KommentarVW = n.kommentarvw  LEFT JOIN `laborparameter` p ON n.abkü = p.abkü AND n.einheit = p.einheit AND l.langtext = p.langtext LEFT JOIN `laborxpgl` gl ON p.id = gl.idpara LEFT JOIN `laborxpneu` n2 ON gl.idxpbez = n2.id  LEFT JOIN `laborxpnb` nb2 ON n2.id = nb2.pid AND nb2.geschlecht IN (3,9,0,IF(na.geschlecht='m',IF(ADDDATE(na.gebdat,INTERVAL 18 YEAR)>CAST(eingang As Date),1,4) ,IF(ADDDATE(na.gebdat,INTERVAL 18 YEAR)>CAST(eingang As Date),2,5))) WHERE wert <> '' AND NOT ISNULL(wert) AND NOT EXISTS(SELECT 0 FROM `laborxus` u LEFT JOIN `laborxwert` w ON u.RefNr = w.RefNr LEFT JOIN `laborxpnb` nb ON w.nbid = nb.id WHERE u.pat_id = n.pat_id AND w.abkü = n.abkü AND w.wert = n.wert AND CAST(n.zeitpunkt As Date) > SUBDATE(CAST(u.Eingang As Date),INTERVAL 3 DAY) AND CAST(n.zeitpunkt As Date) < ADDDATE(CAST(u.Eingang As Date),INTERVAL 6 DAY))"

'Vsql = "SELECT n.Pat_ID Pat_ID,CAST(n.ZeitPunkt As Date) ZeitPunkt,n.FertigStGrad FertigStGrad," & vbCrLf & _
"IF(ISNULL(n2.abkü),n.abkü,n2.abkü) Abkü, n.Abkü abk_ur,l.Langtext Langtext," & vbCrLf & _
"IF(trim(n.Wert) REGEXP '^[0-9]+\\,?[0-9]*$', REPLACE(trim(n.Wert),',','.'),trim(n.Wert)) Wert," & vbCrLf & _
"IF(ISNULL(n2.abkü),n.einheit,n2.einheit) Einheit, n.Einheit Einheit_ur,k.Kommentar Kommentar, " & vbCrLf & _
"IF(ISNULL(n2.abkü),CONCAT(IF(na.geschlecht='m' OR ISNULL(uNw) OR uNw='',uNm,uNw),'-'," & vbCrLf & _
"IF(na.geschlecht='m' OR ISNULL(oNw) OR oNw = '',oNm,oNw)),nb2.nb) NB, " & vbCrLf & _
"CONCAT(IF(na.geschlecht='m' OR ISNULL(uNw) OR uNw='',uNm,uNw),'-',IF(na.geschlecht='m' OR ISNULL(oNw) OR oNw='',oNm,oNw)) NB_ur, " & vbCrLf & _
"IF(ISNULL(n2.abkü),IF(na.geschlecht='m' OR ISNULL(uNw) OR uNw='',uNm,uNw),nb2.uNg) uNg,IF(na.geschlecht='m' OR ISNULL(uNw) OR uNw='',uNm,uNw) uNg_ur, " & vbCrLf & _
"IF(ISNULL(n2.abkü),IF(na.geschlecht='m' OR ISNULL(oNw) OR oNw='',oNm,oNw),nb2.oNg) oNg,IF(na.geschlecht='m' OR ISNULL(oNw) OR oNw='',oNm,oNw) oNg_ur," & _
"'TM' Labor, '' Pfad FROM (" & vbCrLf & _
"`laborneu` n " & vbCrLf & _
"LEFT JOIN `namen` na ON n.pat_id = na.pat_id " & vbCrLf & _
"LEFT JOIN `laborlangtext` l ON l.langtextvw = n.langtextvw) " & _
"LEFT JOIN `laborkommentar` k ON k.KommentarVW = n.kommentarvw  " & vbCrLf & _
"LEFT JOIN `laborparameter` p ON n.abkü = p.abkü AND n.einheit = p.einheit " & vbCrLf & _
"LEFT JOIN `laborypgl` gl ON p.id = gl.idpara " & vbCrLf & _
"LEFT JOIN `laborypneu` n2 ON gl.idypbez = n2.id  " & vbCrLf & _
"LEFT JOIN `laborypnb` nb2 ON n2.id = nb2.pid AND nb2.geschlecht IN (3,9,0,IF(na.geschlecht='m',IF(ADDDATE(na.gebdat,INTERVAL 18 YEAR)>CAST(eingang As Date),1,4) ,IF(ADDDATE(na.gebdat,INTERVAL 18 YEAR)>CAST(eingang As Date),2,5))) " & vbCrLf & _
"WHERE wert <> '' AND NOT ISNULL(wert) AND NOT EXISTS(SELECT 0 FROM `laboryus` u " & vbCrLf & _
    "LEFT JOIN `laborywert` w ON u.id = w.usid " & vbCrLf & _
    "LEFT JOIN `laborypnb` nb ON w.nbid = nb.id WHERE u.pat_id = n.pat_id AND w.abkü = n.abkü AND w.wert = n.wert " & vbCrLf & _
            "AND CAST(n.zeitpunkt As Date) > SUBDATE(CAST(u.Eingang As Date),INTERVAL 3 DAY) AND CAST(n.zeitpunkt As Date) < ADDDATE(CAST(u.Eingang As Date),INTERVAL 6 DAY)) AND (ISNULL(n2.eingang) OR n2.eingang = (SELECT MAX(eingang) FROM laborypnb WHERE pid = n2.id))"

'Vsql = _
"SELECT `n`.`Pat_ID` AS `Pat_ID`,CAST(`n`.`ZeitPunkt` As Date) AS `ZeitPunkt`,`n`.`FertigStGrad` AS `FertigStGrad`,IF(ISNULL(`n2`.`Abkü`),`n`.`Abkü`,`n2`.`Abkü`) AS `Abkü`,`n`.`Abkü` AS `abk_ur`,`l`.`Langtext` AS `Langtext`," & _
"`n`.`Wert` AS `Wert`,IF(ISNULL(`n2`.`Abkü`),`n`.`Einheit`,`n2`.`Einheit`) AS `Einheit`,`n`.`Einheit` AS `Einheit_ur`,`k`.`Kommentar` AS `Kommentar`,IF(ISNULL(`n2`.`Abkü`),CONCAT(IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`uNw`) OR (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`),'-',IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`oNw`) OR (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`)),`nb2`.`NB`) AS `NB`,CONCAT(IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`uNw`) OR (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`),'-',IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`oNw`) OR (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`)) AS `NB_ur`,IF(ISNULL(`n2`.`Abkü`),IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`uNw`) OR (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`),`nb2`.`uNg`) AS `uNg`," & _
"IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`uNw`) OR (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`) AS `uNg_ur`,IF(ISNULL(`n2`.`Abkü`),IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`oNw`) OR (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`),`nb2`.`oNg`) AS `oNg`,IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`oNw`) OR (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`) AS `oNg_ur`,'TM' AS `Labor`, '' `Pfad` " & _
"FROM (((((((`laborneu` `n` " & _
"              LEFT JOIN `namen` `na` on((`n`.`Pat_ID` = `na`.`Pat_ID`))) " & _
"              LEFT JOIN `laborlangtext` `l` on((`l`.`LangtextVW` = `n`.`LangtextVW`))) " & _
"              LEFT JOIN `laborkommentar` `k` on((`k`.`KommentarVW` = `n`.`KommentarVW`))) " & _
"              LEFT JOIN `laborparameter` `p` on(((`n`.`Abkü` = `p`.`Abkü`) AND (`n`.`Einheit` = `p`.`Einheit`)))) " & _
"              LEFT JOIN `laborypgl` `gl` on((`p`.`id` = `gl`.`idpara`))) " & _
"              LEFT JOIN `laborypneu` `n2` on((`gl`.`idxpbez` = `n2`.`id`))) " & _
"              LEFT JOIN `laborypnb` `nb2` on(((`n2`.`id` = `nb2`.`pid`) AND (`nb2`.`Geschlecht` IN (3,9,0,IF((`na`.`Geschlecht` = 'm'),IF(((`na`.`GebDat` + INTERVAL 18 YEAR) > CAST(`nb2`.`Eingang` As Date)),1,4),IF(((`na`.`GebDat` + INTERVAL 18 YEAR) > CAST(`nb2`.`Eingang` As Date)),2,5))))))) " & _
"     LEFT JOIN labor2a l2 ON l2.pat_id = n.pat_id AND l2.abk_ur = n.abkü " & _
"        AND l2.fertigstgrad = n.fertigstgrad " & _
"        AND (l2.wert = n.wert OR l2.kommentar = n.wert) " & _
"       AND (CAST(`n`.`ZeitPunkt` As Date) > (CAST(`l2`.`zeitpunkt` As Date) - INTERVAL 3 DAY)) " & _
"       AND (CAST(`n`.`ZeitPunkt` As Date) < (CAST(`l2`.`zeitpunkt` As Date) + INTERVAL 6 DAY)) " & _
" WHERE (`n`.`Wert` <> '') AND (`n`.`Wert` IS NOT NULL) " & _
"  AND (ISNULL(`nb2`.`Eingang`) OR (`nb2`.`Eingang` = (SELECT MAX(`laborypnb`.`Eingang`) FROM `laborypnb` WHERE (`laborypnb`.`pid` = `n2`.`id`)))) " & _
"  AND ISNULL(l2.pat_id)"

Vsql = _
"SELECT n.id, n.Pat_ID Pat_ID,CAST(n.ZeitPunkt As Date) ZeitPunkt, " & vbCrLf & _
"n.FertigStGrad, " & vbCrLf & _
"n.Abkü, " & vbCrLf & _
"l.Langtext, " & vbCrLf & _
"TRIM(IF(n.Abkü='ALBUM' AND n.Wert='' AND k.Kommentar LIKE 'nicht berechenb%','< 20',IF(TRIM(n.Wert) REGEXP '^[0-9]+\\,?[0-9]*$', REPLACE(n.Wert,',','.'),n.Wert))) Wert," & vbCrLf & _
"n.Einheit," & vbCrLf & _
"IF(ISNULL(k.Kommentar),'',k.Kommentar) Kommentar,nb.normbervw," & vbCrLf & _
"NBm NB," & vbCrLf & _
"uNm `uNg`," & vbCrLf & _
"oNm `oNg`," & vbCrLf & _
"_utf8mb4'TM' COLLATE utf8mb4_german2_ci Labor, _utf8mb4'' COLLATE utf8mb4_german2_ci Pfad, 0 DatID " & vbCrLf & _
", Gruppe, Reihe,1 Qu " & vbCrLf & _
"FROM `laborneu` `n` " & vbCrLf & _
"LEFT JOIN labornormber nb USING (normbervw)" & vbCrLf & _
"              LEFT JOIN `namen` `na` ON `n`.`Pat_ID` = `na`.`Pat_ID` " & vbCrLf & _
"              LEFT JOIN `laborlangtext` `l` ON `l`.`LangtextVW` = `n`.`LangtextVW` " & vbCrLf & _
"              LEFT JOIN `laborkommentar` `k` ON `k`.`KommentarVW` = `n`.`KommentarVW` " & vbCrLf & _
"              LEFT JOIN laborparameter p ON n.Abkü = p.Abkü AND n.Einheit = IF(p.einheit IN ('','\'kA\''),'kA',p.Einheit) " & vbCrLf & _
"                    AND p.NBm = nb.normber" & vbCrLf & _
"                    AND p.id=(SELECT MIN(id) FROM laborparameter WHERE abkü=n.`Abkü` AND IF(einheit IN ('','\'kA\''),'kA',einheit)=n.Einheit AND NBm=nb.normber)" & vbCrLf & _
" WHERE (n.Wert <> '' AND n.Wert IS NOT NULL) OR (k.Kommentar<>'' AND k.Kommentar IS NOT NULL)"
'  OR nb.normber IS NULL OR nb.normber=0

' "CONCAT(IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`uNw`) OR (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`),'-',IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`oNw`) OR (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`)) NB, " & vbCrLf & _
' "IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`uNw`) OR (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`) `uNg`," & vbCrLf & _
' "IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`oNw`) OR (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`) `oNg`," & vbCrLf & _


' "CONCAT(IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`uNw`) OR (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`),'-',IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`oNw`) OR (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`)) `NB_ur`, " & vbCrLf & _
'"IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`uNw`) OR (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`) `uNg_ur`," & vbCrLf & _
'"IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`oNw`) OR (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`) `oNg_ur`," & vbCrLf & _

'"CAST(n.FertigStGrad AS char CHARSET utf8) COLLATE utf8_unicode_ci FertigStGrad, " & vbCrLf & _
"CAST(n.Abkü AS char CHARSET utf8) COLLATE utf8_unicode_ci Abkü, " & vbCrLf & _
"CAST(n.Abkü AS char CHARSET utf8) COLLATE utf8_unicode_ci abk_ur, " & vbCrLf & _
"CAST(l.Langtext AS char CHARSET utf8) COLLATE utf8_unicode_ci Langtext," & vbCrLf & _
"CAST(trim(IF(trim(n.Wert) REGEXP '^[0-9]+\\,?[0-9]*$', REPLACE(n.Wert,',','.'),n.Wert)) AS char CHARSET utf8) COLLATE utf8_unicode_ci Wert," & vbCrLf & _
"CAST(n.Einheit AS char CHARSET utf8) COLLATE utf8_unicode_ci Einheit," & vbCrLf & _
"CAST(n.Einheit AS char CHARSET utf8) COLLATE utf8_unicode_ci Einheit_ur," & vbCrLf & _
"CAST(IF(ISNULL(k.Kommentar),'',k.Kommentar) AS char CHARSET utf8) COLLATE utf8_unicode_ci Kommentar," & vbCrLf & _
"CAST(CONCAT(IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`uNw`) OR (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`),'-',IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`oNw`) OR (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`)) AS char CHARSET utf8) COLLATE utf8_unicode_ci NB, " & vbCrLf & _
"CAST(CONCAT(IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`uNw`) OR (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`),'-',IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`oNw`) OR (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`)) AS char CHARSET utf8) COLLATE utf8_unicode_ci `NB_ur`, " & vbCrLf & _
"CAST(IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`uNw`) OR (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`) AS char CHARSET utf8) COLLATE utf8_unicode_ci `uNg`," & vbCrLf & _
"CAST(IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`uNw`) OR (`p`.`uNw` = '')),`p`.`uNm`,`p`.`uNw`) AS char CHARSET utf8) COLLATE utf8_unicode_ci `uNg_ur`," & vbCrLf & _
"CAST(IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`oNw`) OR (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`) AS char CHARSET utf8) COLLATE utf8_unicode_ci `oNg`," & vbCrLf & _
"CAST(IF(((`na`.`Geschlecht` = 'm') OR ISNULL(`p`.`oNw`) OR (`p`.`oNw` = '')),`p`.`oNm`,`p`.`oNw`) AS char CHARSET utf8) COLLATE utf8_unicode_ci `oNg_ur`," & vbCrLf & _
"_utf8'TM' COLLATE utf8_unicode_ci Labor, _utf8'' COLLATE utf8_unicode_ci Pfad " & vbCrLf & _

Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "formular"
'Vsql = "SELECT `forminhkopf`.FoID,`forminhkopf`.Pat_ID,`forminhkopf`.FID,`forminhkopf`.Form_ID,`forminhkopf`.ZeitPunkt AS ZeitPunkt,`forminhfeld`.Nr AS Nr,`forminhfeld`.FeldNr AS FeldNr,`forminhaltfeld`.Feld,`forminhaltfeldinh`.FeldInh,`formulare`.Form_Abk,`formulare`.FormVorl " & vbcrlf & _
"FROM ((((`forminhfeld` LEFT JOIN `forminhkopf` on((`forminhfeld`.FoID = `forminhkopf`.FoID))) " & vbcrfl & _
"LEFT JOIN `formulare` on((`formulare`.FormID = `forminhkopf`.Form_ID)))  " & vbcrfl & _
"LEFT JOIN `forminhaltfeld` on((`forminhfeld`.FeldVW = `forminhaltfeld`.FeldVW)))  " & vbcrfl & _
"LEFT JOIN `forminhaltfeldinh` on((`forminhfeld`.FeldInhVW = `forminhaltfeldinh`.FeldInhVW)))  " & vbcrfl & _
"ORDER BY `forminhkopf`.FoID;"
Vsql = "SELECT fik.FoID,fik.Pat_ID,fik.FID,fik.Form_ID,fik.ZeitPunkt," & vbCrLf & _
"fi.Nr,fi.FeldNr,fif.Feld,fifi.FeldInh,fo.Form_Abk,fo.FormVorl " & vbCrLf & _
"FROM `forminhfeld` fi" & vbCrLf & _
"LEFT JOIN `forminhkopf` fik ON fik.FoID = fi.FoID" & vbCrLf & _
"LEFT JOIN `formulare` fo ON fo.FormID = fik.Form_ID" & vbCrLf & _
"LEFT JOIN `forminhaltfeld` fif USE KEY (PRIMARY) ON fif.FeldVW = fi.FeldVW" & vbCrLf & _
"LEFT JOIN `forminhaltfeldinh` fifi ON fifi.FeldInhVW = fi.FeldInhVW" & vbCrLf & _
"ORDER BY fik.FoID;"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

' Formular mit Index auswahl für Formularkopf, also nach pat_id auszuwählen (
VN = "formulari"
Vsql = "SELECT `forminhkopf`.FoID AS foid,`forminhkopf`.Pat_ID AS Pat_ID,`forminhkopf`.FID AS FID,`forminhkopf`.Form_ID AS Form_ID,`forminhkopf`.ZeitPunkt AS ZeitPunkt,`forminhfeld`.Nr AS Nr,`forminhfeld`.FeldNr AS FeldNr,`forminhaltfeld`.Feld AS Feld,`forminhaltfeldinh`.FeldInh AS FeldInh,`formulare`.Form_Abk form_abk,`formulare`.FormVorl FormVorl FROM ((((`forminhfeld` LEFT JOIN `forminhkopf` use index(auswahl) on((`forminhfeld`.FoID = `forminhkopf`.FoID))) LEFT JOIN `formulare` on((`formulare`.FormID = `forminhkopf`.Form_ID))) LEFT JOIN `forminhaltfeld` on((`forminhfeld`.FeldVW = `forminhaltfeld`.FeldVW))) LEFT JOIN `forminhaltfeldinh` on((`forminhfeld`.FeldInhVW = `forminhaltfeldinh`.FeldInhVW))) ORDER BY `forminhkopf`.FoID;"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "__kontakttage"
'IF LVobMySQL AND 1 = 0 THEN
' Vsql = "SELECT pat_id,FIRST(zeitpunkt) FROM `eintraege` e WHERE ZeitPunkt BETWEEN cDATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY))¡'-'¡intacc((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1)divmy 3)*3+1¡'-01')) AND cDATE(ADDDATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY))¡'-'¡intacc((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1)divmy 3)*3+1¡'-01'), INTERVAL 3 MONTH)) AND e.Art IN (" & artspezG & ") GROUP BY e.Pat_ID,intaccdatemy(zeitpunkt)"
'ELSE
''SELECT `e`.`Pat_ID` AS `pat_id`,`e`.`ZeitPunkt` AS `zeitpunkt` FROM `quelle1`.`eintraege` `e` WHERE ((`e`.`ZeitPunkt` BETWEEN CONCAT(YEAR((NOW() - INTERVAL 14 DAY)),'-',((((month((NOW() - INTERVAL 14 DAY)) - 1) DIV 3) * 3) + 1),'-01') AND CONCAT((YEAR((NOW() - INTERVAL 14 DAY)) + ROUND((((((month((NOW() - INTERVAL 14 DAY)) - 1) DIV 3) * 3) + 4) / 12),0)),'-',(((((month((NOW() - INTERVAL 14 DAY)) - 1) DIV 3) * 3) + 4) % 12),'-01')) AND (`e`.`Art` in
''('notiz','telef','ni','gstel','gs','rz','ep','bga','tk','APK','wr','ga','tst','cr','ke','hz','mh','ag','ph','pq','er','ds','st','eb','fa','bz','rp','uzu','hypo','colo','aug','beweg','pros','impf','gyn','caro','beruf','ap','mu','rauch','alko','fams','schula','ass','kra','proc','au','GPD','ba','ARCHIE2','gewicht','gewi','rrvgl','bzvgl','bzm','bztp','bks','anal','andm','usal','usdm','doppler','duplex','sono','sd','UKG','Größe','HbA1c','hyper','fuß','keto','wv','ulc','kv','debr','EKG','LZRR','Lufu','lactoset','trop','temp','oGTT','gpt','bmi','urin','taille','hüfte','puls ','GDT','bef'))) GROUP BY `e`.`Pat_ID`,CAST(`e`.`ZeitPunkt` As Date)
' Vsql = "SELECT pat_id,FIRST(zeitpunkt) FROM `eintraege` e WHERE ZeitPunkt BETWEEN cDATE(CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY))¡'-'¡intacc((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1)divmy 3)*3+1¡'-01')) AND cDATE(       (CONCAT(YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY)) + ROUND((intacc((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1)divmy 3)*3+4) / 12)¡'-'¡(intacc((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1)divmy 3)*3+4) mod 12¡'-01')                  )) AND e.Art IN (" & artspezG & ") GROUP BY e.Pat_ID,intaccdatemy(zeitpunkt)"
'END IF
Vsql = "SELECT pat_id, first(zeitpunkt) FROM `eintraege` e WHERE zeitpunkt " & DiesQ() & " AND e.Art IN (" & artspezG & ") GROUP BY e.Pat_ID,intaccdatemy(zeitpunkt)"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


'vsql(5) = "SELECT pat_id,zeitpunkt FROM `eintraege` e WHERE ZeitPunkt BETWEEN STR_TO_DATE(CONCAT('#'¡intacc((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1)divmy 3)*3+1¡'/01/'¡YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY))¡'#'),'#%m/%d/%Y#') AND ADDDATE(STR_TO_DATE(CONCAT('#'¡intacc((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1)divmy 3)*3+1¡'/01/'¡YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY))¡'#'),'#%m/%d/%Y#'), INTERVAL 3 MONTH) AND e.Art IN ('notiz','ni','telef','gs','rz','ga','ag','hz','ts','cr','ep','tst','wr','ek','ph','bga','gstel','rz','ga','ag','hz','ts','ke','uew','rp','sono','duplex','andm','anal')  GROUP BY e.Pat_ID,DATE(zeitpunkt)"
VN = "_kontaktzahl"
Vsql = "SELECT COUNT(0) AS ct,pat_id FROM __kontakttage GROUP BY pat_id;"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

' ca. 3% langsamer als diagview, wohl wegen mehr Spalten
'VN = "diagv"
'Vsql = _
"SELECT " & vbCrLf & _
"  IF(diagsicherheit IN ('G',' '),IF(obdauer,krd,akt),'') gicdko" & vbCrLf & _
", IF(diagsicherheit IN ('G',' '),IF(obdauer,nfc,akt),'') gicdok" & vbCrLf & _
", IF(diagsicherheit IN ('G',' '),IF(obdauer,IF(kori(),krd,nfc),akt),'') gicd" & vbCrLf & _
", IF(diagsicherheit IN ('G',' ','Z'),IF(obdauer,krd,akt),'') gzicdko" & vbCrLf & _
", IF(diagsicherheit IN ('G',' ','Z'),IF(obdauer,nfc,akt),'') gzicdok" & vbCrLf & _
", IF(diagsicherheit IN ('G',' ','Z'),IF(obdauer,IF(kori(),krd,nfc),akt),'') gzicd" & vbCrLf & _
", d.* FROM (" & vbCrLf & _
"SELECT IF(obdauer AND obkasse AND lkasse>qanf(),icd,'') krd" & vbCrLf & _
", IF(Dggel,'',icd) nfc" & vbCrLf & _
", IF(NOT obdauer AND diagdatum<qanf(),'',icd) akt" & vbCrLf & _
", d.* FROM diagnosen d) d"
'Call DtbCreateQueryDef(VN, Vsql)


VN = "diagview"
Vsql = _
"SELECT d.* " & vbCrLf & _
", IF(diagsicherheit IN ('G',' '),IF(obdauer,IF(obkasse AND lkasse>qanf(),icd,''),IF(diagdatum<qanf(),'',icd)),'') gicdko" & vbCrLf & _
", IF(diagsicherheit IN ('G',' '),IF(obdauer,IF(Dggel,'',icd),IF(diagdatum<qanf(),'',icd)),'') gicdok" & vbCrLf & _
", IF(diagsicherheit IN ('G',' '),IF(obdauer,IF(kori(),IF(obkasse AND lkasse>qanf(),icd,''),IF(Dggel,'',icd)),IF(diagdatum<qanf(),'',icd)),'') gicd" & vbCrLf & _
", IF(diagsicherheit IN ('G',' ','Z'),IF(obdauer,IF(obkasse AND lkasse>qanf(),icd,''),IF(diagdatum<qanf(),'',icd)),'') gzicdko" & vbCrLf & _
", IF(diagsicherheit IN ('G',' ','Z'),IF(obdauer,IF(Dggel,'',icd),IF(diagdatum<qanf(),'',icd)),'') gzicdok" & vbCrLf & _
", IF(diagsicherheit IN ('G',' ','Z'),IF(obdauer,IF(kori(),IF(obkasse AND lkasse>qanf(),icd,''),IF(Dggel,'',icd)),IF(diagdatum<qanf(),'',icd)),'') gzicd" & vbCrLf & _
"FROM diagnosen d" & vbCrLf & _
""

' 3,0 s:
' ""

' 5,3 s:
' "GROUP BY pat_id,DiagSicherheit,diagtext,d.diagseite,diagattr,icd,obdauer,intBemerk,AusnBegr,Dggel,obkasse,lkasse,KFdFA"

' 3,4 s:
'"WHERE NOT EXISTS (SELECT 0 FROM diagnosen WHERE id1<d.id1 AND pat_id=d.Pat_id AND diagsicherheit=d.DiagSicherheit AND diagtext=d.diagtext AND diagseite=d.diagseite AND diagattr=d.diagattr AND icd=d.icd AND obdauer=d.obdauer AND intbemerk=d.intBemerk AND ausnbegr=d.AusnBegr AND Dggel=d.Dggel AND obkasse=d.obkasse AND lkasse=d.lkasse AND KFdFA=d.KFdFA)"

' 4,1 s:
'", ROW_NUMBER() OVER (PARTITION BY pat_id,DiagSicherheit,diagtext,d.diagseite,diagattr,icd,obdauer,intBemerk,AusnBegr,Dggel,obkasse,lkasse,KFdFA ORDER BY id1) rang" & vbCrLf & _
'"HAVING rang=1"
Call DtbCreateQueryDef(VN, Vsql)

' "  COALESCE((SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id=n.pat_id AND (art IN (" & artspezG & ") OR art LIKE 'rr%' OR art LIKE 'bz%')),0)" & vbCrLf & _
' "" & vbCrLf & _

VN = "namenleb"
Vsql = _
"SELECT" & vbCrLf & _
" IF(qad=0,qfa,qad) leb" & vbCrLf & _
", i.*" & vbCrLf & _
"FROM (SELECT " & vbCrLf & _
"  COALESCE(IF(aufndat>20040701,aufndat,0),0) qad" & vbCrLf & _
", COALESCE((SELECT MAX(fanf) FROM faelle WHERE pat_id=n.pat_id),0) qfa" & vbCrLf & _
", n.*" & vbCrLf & _
"FROM namen n) i;" & vbCrLf & _
""
Call DtbCreateQueryDef(VN, Vsql)

VN = "namenlb"
Vsql = _
"SELECT" & vbCrLf & _
" IF(ge=0,leb,ge) lbeh" & vbCrLf & _
" , i.*" & vbCrLf & _
"FROM (SELECT " & vbCrLf & _
" GREATEST(" & vbCrLf & _
"  COALESCE((SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id=n.pat_id AND (art IN (" & artspezG & ") OR art LIKE 'rr%' OR art LIKE 'bz%')),0)" & vbCrLf & _
" ,COALESCE((SELECT MAX(zeitpunkt) FROM medplan WHERE pat_id=n.pat_id),0)" & vbCrLf & _
" ) ge" & vbCrLf & _
", n.*" & vbCrLf & _
"FROM namenleb n) i;" & vbCrLf & _
""
Call DtbCreateQueryDef(VN, Vsql)


' diagview mit eindeutigen Datensätzen
VN = "diageview"
Vsql = _
"SELECT d.* " & vbCrLf & _
"FROM diagview d" & vbCrLf & _
"WHERE NOT EXISTS (SELECT 0 FROM diagnosen WHERE id1<d.id1 AND pat_id=d.Pat_id AND diagsicherheit=d.DiagSicherheit AND diagtext=d.diagtext AND diagseite=d.diagseite AND diagattr=d.diagattr AND icd=d.icd AND obdauer=d.obdauer AND intbemerk=d.intBemerk AND ausnbegr=d.AusnBegr AND Dggel=d.Dggel AND obkasse=d.obkasse AND lkasse=d.lkasse AND KFdFA=d.KFdFA)"
Call DtbCreateQueryDef(VN, Vsql)

If LVobMySQL Then
 VN = "aktfaelle"
  Vsql = "SELECT f.pat_id pid, notiz, stru.leistung stru, chron.leistung chron, kt.ct kt, ebm.leistung verspau " & vbCrLf & _
 ",(SELECT icd FROM diagview d WHERE d.pat_id=f.pat_id AND (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd LIKE 'O24%' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND d.diagdatum BETWEEN qbegs(f.quartal) AND qends(f.quartal))) ORDER BY icd LIMIT 1) icd " & vbCrLf & _
 ", f.*, k.id, k.name kname, kateg, anzahlik, anzahlktug, gültigvon, gültigbis, go, kurzname " & vbCrLf & _
 "FROM ((`faelle` f " & vbCrLf & _
 "LEFT JOIN `kassenliste` k ON f.vknr = k.vknr AND f.ik = k.ik) " & vbCrLf & _
 "LEFT JOIN `leistungen` ebm ON f.fid = ebm.fid AND (ebm.leistung LIKE '031%' OR ebm.leistung LIKE '01210') LEFT JOIN `leistungen` chron ON f.fid = chron.fid AND (chron.leistung = '03212') " & vbCrLf & _
 "LEFT JOIN `leistungen` stru ON f.fid = stru.fid AND (stru.leistung LIKE '973%') " & vbCrLf & _
 "LEFT JOIN _kontaktzahl kt ON kt.pat_id = f.pat_id " & vbCrLf & _
 "LEFT JOIN `namen` n ON n.pat_id = f.pat_id) " & vbCrLf & _
 "WHERE schgr <> '90' AND NOT goäkatnr IN ('40','41') AND n.nachname <> 'Bereitschaftsdienst' AND quartal = (SELECTmy  CONCAT(intacc(((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1) divmy 3) + 1) ¡ YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY)) ) AS `lq`) " & vbCrLf & _
 "GROUP BY f.fid ORDER BY pid, schgr, icd"
 Call DtbCreateQueryDef(VN, Vsql)
 vz = vz + 1
End If
'  "LEFT JOIN `leistungen` ebm ON f.fid = ebm.fid AND (ebm.leistung LIKE _utf8mb4'031%' COLLATE utf8mb4_german2_ci OR ebm.leistung LIKE _utf8mb4'01210' COLLATE utf8mb4_german2_ci) LEFT JOIN `leistungen` chron ON f.fid = chron.fid AND (chron.leistung = _utf8mb4'03212' COLLATE utf8mb4_german2_ci) " & vbCrLf & _
 "LEFT JOIN `leistungen` stru ON f.fid = stru.fid AND (stru.leistung LIKE _utf8mb4'973%' COLLATE utf8mb4_german2_ci) " & vbCrLf & _
 "LEFT JOIN _kontaktzahl kt ON kt.pat_id = f.pat_id " & vbCrLf & _
 "LEFT JOIN `namen` n ON n.pat_id = f.pat_id) " & vbCrLf & _
 "WHERE schgr <> '90' AND NOT goäkatnr IN ('40','41') AND n.nachname <> 'Bereitschaftsdienst' AND quartal = (SELECTmy  CONCAT(intacc(((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1) divmy 3) + 1) ¡ YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY)) COLLATE utf8mb4_german2_ci) AS `lq`) " & vbCrLf & _


If LVobMySQL Then

 VN = "aktf"
 Vsql = "SELECT pat_id,fid,schgr,vknr,ik,goäkatnr,lanrid,quartal " & vbCrLf & _
        "FROM `faelle` " & vbCrLf & _
        "WHERE schgr <> '90' AND NOT goäkatnr IN ('40','41') " & vbCrLf & _
        "AND nachname <> 'Bereitschaftsdienst' " & vbCrLf & _
        "AND quartal = (SELECTmy  CONCAT(intacc(((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1) divmy 3) + 1) ¡ YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY)) ) AS `lq`) " & _
           "ORDER BY pat_id, fid DESC, schgr;"
 Call DtbCreateQueryDef(VN, Vsql)
 vz = vz + 1
'        "AND nachname <> 'Bereitschaftsdienst' " & vbCrLf & _
        "AND quartal = (SELECTmy  CONCAT(intacc(((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1) divmy 3) + 1) ¡ YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY))  COLLATE utf8mb4_german2_ci) AS `lq`) " & _


 ' aktuelle Faelle verschieden, schnell (ohne Krankenkassenkategorie)
 VN = "aktfvs"
' Vsql = "SELECT pat_id,fid,schgr,vknr,ik,lanrid " & vbCrLf & _
        "FROM `faelle` " & vbCrLf & _
        "WHERE schgr <> '90' AND NOT goäkatnr IN ('40','41') " & vbCrLf & _
        "AND nachname <> 'Bereitschaftsdienst' COLLATE utf8mb4_german2_ci " & vbCrLf & _
        "AND quartal = (SELECTmy  CONCAT(intacc(((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1) divmy 3) + 1) ¡ YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY))  COLLATE utf8mb4_german2_ci) AS `lq`) " & _
        "GROUP BY pat_id " & vbCrLf & _
        "ORDER BY pat_id, fid DESC, schgr;"
 Vsql = "SELECT * FROM aktf f " & vbCrLf & _
        "WHERE vknr<>71800" & vbCrLf & _
        "GROUP BY pat_id " & vbCrLf & _
        "ORDER BY pat_id, fid DESC, schgr;"
 Call DtbCreateQueryDef(VN, Vsql)
 vz = vz + 1
   
 VN = "aktfvmi" ' mit Impffaellen bei Privatpatienten (Bayerisches Landesinstitut für Gesundheit)
' Vsql = "SELECT pat_id, fid,schgr,goäkatnr,vknr FROM `faelle` WHERE schgr <> '90' AND NOT goäkatnr IN (40,41) AND nachname <> 'Bereitschaftsdienst' COLLATE utf8mb4_german2_ci AND quartal = (SELECTmy  CONCAT(intacc(((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1) divmy 3) + 1) ¡ YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY)) COLLATE utf8mb4_german2_ci) AS `lq`) " & _
'           "GROUP BY pat_id ORDER BY pat_id;"
' 30.10.16: kürzer ginge auch, Performance getestet aktuell gleich:
' Vsql = "SELECT pat_id,fid,schgr,goäkatnr,vknr,ik,(SELECT MAX(kateg) FROM kassenliste WHERE vknr=f.vknr AND ik=f.ik) kateg, lanrid " & vbCrLf & _
        "FROM `faelle` f " & vbCrLf & _
        "WHERE schgr <> '90' AND NOT goäkatnr IN ('40','41') AND nachname <> 'Bereitschaftsdienst' COLLATE utf8mb4_german2_ci " & _
        "AND quartal=(SELECTmy  CONCAT(intacc(((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1) divmy 3) + 1) ¡ YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY))  COLLATE utf8mb4_german2_ci) AS `lq`) " & _
        "GROUP BY pat_id ORDER BY pat_id;"
 Vsql = "SELECT f.*,(SELECT MAX(kateg) FROM kassenliste WHERE vknr=f.vknr AND ik=f.ik) kateg" & vbCrLf & _
        "FROM aktf f" & vbCrLf & _
        "WHERE quartal = quartal(SUBDATE(NOW(), INTERVAL " & FristS & " DAY)) " & vbCrLf & _
        "GROUP BY pat_id ORDER BY pat_id;"
 Call DtbCreateQueryDef(VN, Vsql)
 vz = vz + 1
 
 VN = "aktfv" ' ohne Impffaelle bei Privatpatienten (Bayerisches Landesinstitut für Gesundheit)
' Vsql = "SELECT pat_id, fid,schgr,goäkatnr,vknr FROM `faelle` WHERE schgr <> '90' AND NOT goäkatnr IN (40,41) AND nachname <> 'Bereitschaftsdienst' COLLATE utf8mb4_german2_ci AND quartal = (SELECTmy  CONCAT(intacc(((month(SUBDATE(NOW(),INTERVAL " & frist & " DAY))-1) divmy 3) + 1) ¡ YEAR(SUBDATE(NOW(),INTERVAL " & frist & " DAY)) COLLATE utf8mb4_german2_ci) AS `lq`) " & _
'           "GROUP BY pat_id ORDER BY pat_id;"
' 30.10.16: kürzer ginge auch, Performance getestet aktuell gleich:
 Vsql = "SELECT * FROM aktfvmi WHERE vknr<>71800;"
'         "AND quartal = quartal(SUBDATE(NOW(), INTERVAL " & fristS & " DAY)) " & _

 Call DtbCreateQueryDef(VN, Vsql)
 vz = vz + 1
    
    sql = "DROP FUNCTION IF EXISTS `qanf`"
    myEFrag (sql)
    sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `qanf`() RETURNS DATE" & Chr$(13) & _
    "NO SQL DETERMINISTIC" & vbCrLf & _
    "BEGIN " & Chr$(13) & _
    " RETURN DATE(CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY),'-',(QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+1,'-01'));" & Chr$(13) & _
    "END "
    myEFrag (sql)

    sql = "DROP FUNCTION IF EXISTS `qend`"
    myEFrag (sql)
    sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `qend`() RETURNS DATETIME" & Chr$(13) & _
    "NO SQL DETERMINISTIC" & vbCrLf & _
    "BEGIN " & vbCrLf & _
    " RETURN CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY)+ quarter(NOW()-INTERVAL " & frist & " DAY) div 4 ,'-',((QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+4) mod 12,'-01')-INTERVAL 1 second;" & Chr$(13) & _
    "END "
    myEFrag (sql)
    
    ' Kodierrichtlinien eingeschaltet
    sql = "DROP FUNCTION IF EXISTS `kori`"
    myEFrag (sql)
    sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `kori`() RETURNS BIT" & Chr$(13) & _
    "CONTAINS SQL DETERMINISTIC" & vbCrLf & _
    "BEGIN " & vbCrLf & _
    " RETURN COALESCE((SELECT ob FROM koricht ORDER BY datum DESC LIMIT 1),0); " & vbCrLf & _
    "END "
    myEFrag (sql)
    
#If False Then
' Folgendes hülfe nur ein bißchen, wenn nicht als DETERMINISTIC deklariert
    " IF ISNULL(@qende) THEN SET @qende:= CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY)+ quarter(NOW()-INTERVAL " & frist & " DAY) div 4 ,'-',((QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+4) mod 12,'-01')-INTERVAL 1 SECOND; END IF;" & vbCrLf & _
    " RETURN @qende;" & vbCrLf & _

' Folgendes verlangsamt die Funktion nochmal doppelt:
    " DECLARE jt DATE DEFAULT SUBDATE(NOW()," & frist & ");" & vbCrLf & _
    " DECLARE j VARCHAR(4) DEFAULT LEFT(jt,4);" & vbCrLf & _
    " DECLARE m VARCHAR(5) DEFAULT MID(jt,6,5);" & vbCrLf & _
    " DECLARE en VARCHAR(4) DEFAULT CASE WHEN m<'04-01' THEN '0331' WHEN m<'07-01' THEN '0630' WHEN m<'10-01' THEN '0930' ELSE '1231' END;" & vbCrLf & _
    " RETURN CONCAT(j,en,'235959');" & vbCrLf & _

#End If
' aktfaelle verschieden mit Therapieart
 VN = "aktfvmta"
 Vsql = _
 "SELECT CASE WHEN dicd LIKE 'E10%' THEN '1' WHEN dicd LIKE 'E11%' THEN '2' WHEN dicd = 'O24.4' THEN 'g' ELSE '-' END dtyp, i.* FROM (" & vbCrLf & _
 " SELECT COALESCE(t.therart,'Diät') mta, rang(COALESCE(t.therart,'Diät')) rang, t.zp tzp " & vbCrLf & _
 ", (SELECT MAX(icd) FROM diagview d WHERE pat_id=f.pat_id AND ((gicd RLIKE '^E1[0-4]\.' AND obdauer<>0) OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit IN ('G',' ') AND obdauer=0 AND diagdatum BETWEEN qanf() AND qend()))) dicd " & vbCrLf & _
 ", f.* " & vbCrLf & _
 " FROM aktfv f " & vbCrLf & _
 " LEFT JOIN therarten t ON t.pat_id=f.pat_id AND t.zp BETWEEN COALESCE((SELECT MAX(zp) FROM therarten WHERE pat_id=f.pat_id AND zp<qanf()),19000101) AND qend() " & vbCrLf & _
 " GROUP BY f.pat_id, rang(t.therart) DESC " & vbCrLf & _
 ") i " & vbCrLf & _
 "GROUP BY pat_id"
 Call DtbCreateQueryDef(VN, Vsql)
 vz = vz + 1

End If

'vn = "lfaelle"
'vsql = Vsql(1)
'call DtbCreateQueryDef(vn,vsql)
'
'
'
VN = "_faellenachschgr"
Vsql = "SELECT * FROM `faelle` ORDER BY schgr"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "faelleverschieden"
If LVobMySQL Then
 Vsql = "SELECT * FROM `_faellenachschgr` AS f GROUP BY pat_id,quartal"
Else
 Vsql = "SELECT `faelle`.* FROM (SELECT FIRST(fid) AS mfid, pat_id, quartal FROM `_faellenachschgr` f GROUP BY f.Pat_ID, f.Quartal) AS f INNER JOIN `faelle` ON f.mfid = `faelle`.fid"
End If
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "aktfaellev"
Vsql = "SELECT * FROM `aktfaelle` GROUP BY pat_id"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "lfaellev"
If LVobMySQL Then
' vsql = "SELECT * FROM `lfaelle` GROUP BY pat_id"
 Vsql = "SELECT f.pat_id, f.fid, f.schgr, f.bhfb, f.ik, f.vknr " & vbCrLf & _
        "FROM `_lfaelle` l LEFT JOIN `faelle` f ON f.pat_id = l.pid AND f.bhfb = l.mbhfb GROUP BY pat_id"
Else
' vsql = "SELECT f.* FROM `faelle` f right JOIN (SELECT pat_id, MIN(fid) AS mfid FROM `lfaelle` GROUP BY pat_id) AS p ON f.pat_id = p.pat_id AND f.fid = p.mfid"
 Vsql = "SELECT f.pat_id, FIRST(f.fid), FIRST(f.schgr), FIRST(f.bhfb), FIRST(f.ik), FIRST(f.vknr) " & vbCrLf & _
        "FROM `_lfaelle` l LEFT JOIN `faelle` f ON f.pat_id = l.pid AND f.bhfb = l.mbhfb GROUP BY pat_id"
End If
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "DMPInkonsistenzen"
If 1 = 0 Then
Vsql = "SELECT f.quartal, COUNT(0) AS ct, o.pat_id, o.fid, o.zeitpunkt, feldinh, form_abk FROM`formular`o LEFT JOIN `faelle` f USING (fid)" & _
           "WHERE form_abk LIKE 'dmp%' AND " & _
           "feld = 'Nachname' AND " & _
           "(`f`.`Quartal` = " & _
           "(SELECT CONCAT((((month((NOW() - INTERVAL 23 DAY)) - 1) DIV 3) + 1),YEAR((NOW() - INTERVAL 23 DAY))) )) " & _
           "GROUP BY pat_id " & _
           "HAVING ct > 1;"
End If
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "faelleverschiedenneu"
If LVobMySQL Then
 Vsql = "SELECT *,(SELECT MIN(fanf) FROM `faelle` f1 WHERE pat_id = f.pat_id AND f1.fanf < f.fanf) AS erst  FROM `_faellenachschgr` AS f GROUP BY pat_id,quartal"
Else
 Vsql = "SELECT (SELECT MIN(fanf) FROM `faelle` f1 WHERE pat_id = f0.pat_id AND f1.fanf < ffanf) AS erst, `faelle`.* FROM (SELECT FIRST(fid) AS mfid, FIRST(fanf) AS ffanf, pat_id, quartal FROM `_faellenachschgr` f GROUP BY f.Pat_ID, f.Quartal) AS f0 INNER JOIN `faelle` ON f0.mfid = `faelle`.fid"
End If
'IF LVobMySQL THEN
' vsql(16) = "SELECT mp.* FROM (SELECT pat_id, mpnr FROM (SELECT mpi.pat_id, mpi.mpnr, mpi.zeitpunkt FROM (SELECT pat_id, MAX(zeitpunkt) AS zp FROM `medplan` mp WHERE zeitpunkt <= now() GROUP BY pat_id) AS mpl LEFT JOIN `medplan` mpi ON mpl.pat_id = mpi.pat_id AND mpl.zp = mpi.zeitpunkt GROUP BY mpl.pat_id, mpl.zp, mpnr ORDER BY  mpl.pat_id, mpl.zp, mpnr DESC) AS mpii GROUP BY pat_id, zeitpunkt) AS innen LEFT JOIN `medplan` mp ON innen.pat_id = mp.pat_id AND innen.mpnr = mp.mpnr"
'END IF
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "_f1"
Vsql = "SELECT bhfb, pat_id FROM `faelle` ORDER BY pat_id,bhfb DESC,schgr"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "letztefaelleverschieden"
Vsql = "SELECT * FROM `_f1` AS i GROUP BY pat_id   "
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

' diagnosenview für die Medizin
VN = "diagmed"
Vsql = _
"SELECT * FROM diagview d" & vbCrLf & _
"WHERE ISNULL(lkasse) OR lkasse=(SELECT MAX(lkasse) FROM diagnosen WHERE pat_id=d.pat_id AND icd=d.icd AND diagtext=d.diagtext AND diagsicherheit=d.diagsicherheit AND diagseite=d.diagseite AND diagattr=d.diagattr AND intbemerk=d.intbemerk AND obdauer=d.obdauer AND ausnbegr=d.ausnbegr)" & vbCrLf & _
"GROUP BY pat_id, icd, diagtext, diagsicherheit, diagseite, diagattr, intbemerk, obdauer, ausnbegr, lkasse"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

' diagnosenview für die Abrechnung, zumindest so lange Diagnosen noch doppelt importiert
VN = "diagabr"
Vsql = _
"SELECT d.* FROM diagview d -- LEFT JOIN koricht k ON datum=(SELECT MAX(datum) FROM koricht)" & vbCrLf & _
"WHERE ISNULL(lkasse) OR lkasse=(SELECT MAX(lkasse) FROM diagview WHERE pat_id=d.pat_id AND icd=d.icd AND diagtext=d.diagtext AND diagsicherheit=d.diagsicherheit AND diagseite=d.diagseite AND diagattr=d.diagattr AND intbemerk=d.intbemerk AND obdauer=d.obdauer AND ausnbegr=d.ausnbegr)" & vbCrLf & _
"GROUP BY pat_id, icd, diagtext, diagsicherheit, diagseite, diagattr, intbemerk, obdauer, ausnbegr, lkasse"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


'VN = "_lHbA1c"
'Vsql = "SELECT f.pat_id, IF(MAX(l2.zeitpunkt) > MAX(l1.zeitpunkt) OR ISNULL(MAX(l1.zeitpunkt)),l2.zeitpunkt,l1.zeitpunkt) lzp, IF(MAX(l2.zeitpunkt) > MAX(l1.zeitpunkt) OR ISNULL(MAX(l1.zeitpunkt)),CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)),CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2))) letzter " & _
       "FROM `namen` f " & _
       "LEFT JOIN `labor2a` l2 ON f.pat_id = l2.pat_id AND l2.abk_ur RLIKE '^hba[1c]' AND l2.einheit = '%' AND l2.zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date) " & _
       "LEFT JOIN `labor1a` l1 ON f.pat_id = l1.pat_id AND l1.abk_ur RLIKE '^hba[1c]' AND l1.einheit = '%' AND l1.zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date) " & _
       "WHERE NOT ISNULL(f.pat_id) GROUP BY pat_id, l2.zeitpunkt ORDER BY f.pat_id, l2.zeitpunkt DESC, l1.zeitpunkt DESC"
'Call DtbCreateQueryDef(VN, Vsql)
'vz = vz + 1

'VN = "lHbA1c"
'Vsql = "SELECT pat_id, lzp, letzter FROM `_lHbA1c` GROUP BY pat_id"
'Call DtbCreateQueryDef(VN, Vsql)
'vz = vz + 1
VN = "lHbA1c"
Vsql = "SELECT n.pat_id, IF(l2.eingang>l1.zeitpunkt OR ISNULL(l1.zeitpunkt),l2.eingang,l1.zeitpunkt) lzp, " & vbCrLf & _
       "CONVERT(REPLACE(CONCAT('0',trim(IF(l2.eingang>l1.zeitpunkt OR ISNULL(l1.zeitpunkt),l2.wert,l1.wert))),',','.'),DECIMAL(9,2)) letzter " & vbCrLf & _
       "FROM namen n " & vbCrLf & _
  "LEFT JOIN (SELECT eingang, wert, abkü, pat_id FROM laboryus u INNER JOIN laborywert w ON w.usid = u.id AND w.abkü RLIKE '^hba[1c]' " & vbCrLf & _
       "AND u.eingang = (SELECT MAX(eingang) FROM laborywert wi INNER JOIN laboryus ui ON wi.usid = ui.id WHERE abkü RLIKE '^hba[1c]' AND pat_id=u.pat_id)) l2 ON l2.pat_id = n.pat_id " & vbCrLf & _
  "LEFT JOIN labor1a l1 ON n.pat_id = l1.pat_id AND l1.abkü RLIKE '^hba[1c]' " & vbCrLf & _
       "AND l1.zeitpunkt = (SELECT MAX(zeitpunkt) FROM labor1a WHERE pat_id = n.pat_id AND abkü RLIKE '^hba[1c]') " & vbCrLf & _
  "GROUP BY pat_id"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "lKrea"
Dim krit$
krit = " abkü IN ('creat','krea02','krea','krea02','kres') AND einheit='mg/dl' "
Vsql = "SELECT n.pat_id, IF(l2.zeitpunkt>l1.zeitpunkt OR ISNULL(l1.zeitpunkt),l2.zeitpunkt,l1.zeitpunkt) lzp" & vbCrLf & _
",IF(l2.zeitpunkt>l1.zeitpunkt OR ISNULL(l1.zeitpunkt),l2.wert,l2.wert) letzter FROM namen n " & _
  "LEFT JOIN (SELECT eingang zeitpunkt, wert, abkü, pat_id FROM laboryus u INNER JOIN laborywert w ON w.usid = u.id AND " & krit & vbCrLf & _
       "AND u.eingang = (SELECT MAX(eingang) FROM laborywert wi INNER JOIN laboryus ui ON wi.usid = ui.id WHERE " & krit & " AND pat_id=u.pat_id)) l2 ON l2.pat_id = n.pat_id " & vbCrLf & _
  "LEFT JOIN labor1a l1 ON n.pat_id = l1.pat_id AND l1.abkü IN ('creat','krea02','krea','krea02','kres') AND l1.einheit='mg/dl' AND l1.zeitpunkt = (SELECT MAX(zeitpunkt) FROM labor1a WHERE pat_id = n.pat_id and" & krit & ") " & _
  "GROUP BY pat_id"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1
'   "LEFT JOIN labor2a l2 ON n.pat_id = l2.pat_id AND l2.abkü IN ('creat','krea02','krea','krea02','kres') AND l2.einheit='mg/dl' AND l2.zeitpunkt = (SELECT MAX(zeitpunkt) FROM labor2a WHERE pat_id = n.pat_id and" & krit & ") " & _

VN = "lGFR"
krit = " abkü RLIKE '^gfr|^gfck|^creacl|^m' "
Vsql = "SELECT n.pat_id, IF(l2.zeitpunkt>l1.zeitpunkt OR ISNULL(l1.zeitpunkt),l2.zeitpunkt,l1.zeitpunkt) lzp,  CONVERT(REPLACE(CONCAT('0',trim(IF(l2.zeitpunkt>l1.zeitpunkt OR ISNULL(l1.zeitpunkt),l2.wert,l1.wert))),',','.'),DECIMAL(9,2)) letzter " & vbCrLf & _
  "FROM namen n " & _
  "LEFT JOIN (SELECT eingang zeitpunkt, wert, abkü, pat_id FROM laboryus u INNER JOIN laborywert w ON w.usid = u.id AND " & krit & vbCrLf & _
       "AND u.eingang = (SELECT MAX(eingang) FROM laborywert wi INNER JOIN laboryus ui ON wi.usid = ui.id WHERE " & krit & " AND pat_id=u.pat_id)) l2 ON l2.pat_id = n.pat_id " & vbCrLf & _
  "LEFT JOIN labor1a l1 ON n.pat_id = l1.pat_id AND l1.abkü RLIKE '^gfr|^gfck|^creacl' AND l1.zeitpunkt = (SELECT MAX(zeitpunkt) FROM labor1a WHERE pat_id = n.pat_id AND abkü RLIKE '^gfr|^gfck|^creacl') " & _
  "GROUP BY pat_id"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1
'   "LEFT JOIN labor2a l2 ON n.pat_id = l2.pat_id AND l2.abkü RLIKE '^gfr|^gfck|^creacl|^m' AND l2.zeitpunkt = (SELECT MAX(zeitpunkt) FROM labor2a WHERE pat_id = n.pat_id AND abkü RLIKE '^gfr|^gfck|^creacl') " & _


VN = "DMP-HbA1c-Statistik"
Vsql = "SELECT SUM(IF(letzter>=8.5,1,0)) `HbA1c>8,5`, SUM(IF(NOT ISNULL(letzter),1,0)) `HbA1c vorh`,COUNT(0) `Zahl akt.Fälle`, ROUND(SUM(IF(letzter>=8.5,1,0))/SUM(IF(NOT ISNULL(letzter),1,0))*100,2) Anteil_an_Vorh, ROUND(SUM(IF(letzter>=8.5,1,0))/COUNT(0)*100,2) Anteil_an_Fallzahl " & vbCrLf & _
       "FROM `aktfvs` f " & vbCrLf & _
       "LEFT JOIN namen n ON f.pat_id = n.pat_id " & vbCrLf & _
       "LEFT JOIN lHbA1c h ON f.pat_id = h.pat_id " & vbCrLf & _
       "WHERE dmpklass =3"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "_HbA1cMinMax" ' 4 x gestrichen: AND einheit = '%'
'Vsql = "SELECT n.pat_id," & _
       "(SELECT MAX(CONVERT(REPLACE(CONCAT('0',trim(ln.wert)),',','.'),DECIMAL(9,2))) FROM laborneu ln WHERE ln.pat_id = n.pat_id AND ln.abkü RLIKE '^hba[1c]' AND CONVERT(REPLACE(CONCAT('0',trim(ln.wert)),',','.'),DECIMAL(9,2)) <21 AND ln.wert RLIKE '^[0-9 .,]*$' AND zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date)) max1," & _
       "(SELECT MAX(CONVERT(REPLACE(CONCAT('0',trim(w.wert)),',','.'),DECIMAL(9,2))) FROM laborywert w LEFT JOIN laboryus u ON w.usid = u.id WHERE u.pat_id = n.pat_id AND w.abkü RLIKE '^hba[1c]' AND CONVERT(REPLACE(CONCAT('0',trim(w.wert)),',','.'),DECIMAL(9,2))<21 AND w.wert RLIKE '^[0-9 .,]*$' AND eingang > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date)) max2, " & _
       "(SELECT MIN(CONVERT(REPLACE(CONCAT('0',trim(ln.wert)),',','.'),DECIMAL(9,2))) FROM laborneu ln WHERE ln.pat_id = n.pat_id AND ln.abkü RLIKE '^hba[1c]' AND CONVERT(REPLACE(CONCAT('0',trim(ln.wert)),',','.'),DECIMAL(9,2)) <21 AND ln.wert RLIKE '^[0-9 .,]*$' AND zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date)) min1," & _
       "(SELECT MIN(CONVERT(REPLACE(CONCAT('0',trim(w.wert)),',','.'),DECIMAL(9,2))) FROM laborywert w LEFT JOIN laboryus u ON w.usid = u.id WHERE u.pat_id = n.pat_id AND w.abkü RLIKE '^hba[1c]' AND CONVERT(REPLACE(CONCAT('0',trim(w.wert)),',','.'),DECIMAL(9,2))<21 AND w.wert RLIKE '^[0-9 .,]*$' AND eingang > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date)) min2 " & _
       "FROM namen n"
Vsql = "SELECT pat_id" & vbCrLf & _
    ",(SELECT MIN(CAST(wert AS DECIMAL(9,2))) FROM labor1a WHERE pat_id = n.pat_id AND abkü RLIKE '^hba[1c]' AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND wert<21) min1" & vbCrLf & _
    ",(SELECT MIN(CAST(wert AS DECIMAL(9,2))) FROM laborywert w INNER JOIN laboryus u ON w.usid = u.id WHERE pat_id = n.pat_id AND abkü RLIKE '^hba[1c]' AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND wert<21) min2" & vbCrLf & _
    ",(SELECT MAX(CAST(wert AS DECIMAL(9,2))) FROM labor1a WHERE pat_id = n.pat_id AND abkü RLIKE '^hba[1c]' AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND wert<21) max1" & vbCrLf & _
    ",(SELECT MAX(CAST(wert AS DECIMAL(9,2))) FROM laborywert w INNER JOIN laboryus u ON w.usid = u.id WHERE pat_id = n.pat_id AND abkü RLIKE '^hba[1c]' AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND wert<21) max2" & vbCrLf & _
    "FROM namen n"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "HbA1cMinMax"
Vsql = "SELECT h.pat_id," & _
       "IF(max1>max2 OR ISNULL(max2),max1,max2) max," & _
       "IF(min1<min2 OR ISNULL(min2),min1,min2) min " & _
       "FROM _HbA1cMinMax h"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "_maxHbA1c" ' 2 x gestrichen: AND ln.einheit = '%'
'Vsql = "SELECT n.pat_id," & _
       "(SELECT MAX(CONVERT(REPLACE(CONCAT('0',trim(ln.wert)),',','.'),DECIMAL(9,2))) FROM laborneu ln WHERE ln.pat_id = n.pat_id AND ln.abkü RLIKE '^hba[1c]' AND CONVERT(REPLACE(CONCAT('0',trim(ln.wert)),',','.'),DECIMAL(9,2)) <21 AND ln.wert RLIKE '^[0-9 .,]*$') w1," & _
       "(SELECT MAX(CONVERT(REPLACE(CONCAT('0',trim(w.wert)),',','.'),DECIMAL(9,2))) FROM laborywert w LEFT JOIN laboryus u ON w.usid = u.id WHERE u.pat_id = n.pat_id AND w.abkü RLIKE '^hba[1c]' AND CONVERT(REPLACE(CONCAT('0',trim(w.wert)),',','.'),DECIMAL(9,2))<21 AND w.wert RLIKE '^[0-9 .,]*$') w2 " & _
       "FROM namen n"
Vsql = "SELECT n.pat_id" & _
       ",(SELECT MAX(CAST(wert AS DECIMAL(9,2))) FROM labor1a WHERE pat_id = n.pat_id AND abkü RLIKE '^hba[1c]' AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND (wert<20 OR (wert<21 AND einheit='%'))) max1" & vbCrLf & _
       ",(SELECT MAX(CAST(wert AS DECIMAL(9,2))) FROM laborywert w INNER JOIN laboryus u ON w.usid = u.id WHERE pat_id = n.pat_id AND abkü RLIKE '^hba[1c]' AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND (wert<20 OR (wert<21 AND einheit='%'))) max2" & vbCrLf & _
       "FROM namen n"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "maxHbA1c"
Vsql = "SELECT pat_id, IF(max1>max2 OR ISNULL(max2),max1,max2) Wert FROM _maxHbA1c"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "_maxGluc"
'Vsql = "SELECT n.pat_id," & _
       "(SELECT MAX(CONVERT(REPLACE(CONCAT('0',trim(ln.wert)),',','.'),DECIMAL(9,2))) FROM laborneu ln WHERE ln.pat_id = n.pat_id AND ln.abkü RLIKE '^glu' AND ln.einheit = 'mg/dl') w1," & _
       "(SELECT MAX(CONVERT(REPLACE(CONCAT('0',trim(w.wert)),',','.'),DECIMAL(9,2))) FROM laborywert w LEFT JOIN laboryus u ON w.usid = u.id WHERE u.pat_id = n.pat_id AND w.abkü RLIKE '^glu' AND w.einheit = 'mg/dl') w2 " & _
       "FROM namen n"
Vsql = "SELECT n.pat_id" & _
       ",(SELECT MAX(CAST(wert AS DECIMAL(9,2))) FROM labor1a WHERE pat_id = n.pat_id AND abkü RLIKE '^glu' AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND einheit='mg/dl') max1" & vbCrLf & _
       ",(SELECT MAX(CAST(wert AS DECIMAL(9,2))) FROM laborywert w INNER JOIN laboryus u ON w.usid = u.id WHERE pat_id = n.pat_id AND abkü RLIKE '^glu' AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND einheit='mg/dl') max2" & vbCrLf & _
       "FROM namen n"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "maxGluc"
Vsql = "SELECT pat_id, IF(max1>max2 OR ISNULL(max2),max1,max2) Wert FROM _maxGluc"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "_AlbCreMinMax"
'Vsql = "SELECT n.pat_id, MAX(IF(CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)) > CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)) OR ISNULL(l1.wert), CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)), CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)))) max, MIN(IF(CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)) < CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)) OR ISNULL(l1.wert),CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)),CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)))) min " & _
       "FROM `namen` n " & _
       "LEFT JOIN `labor2a` l2 ON n.pat_id = l2.pat_id AND l2.abk_ur IN ('albcre','albq','album','albup') AND l2.wert REGEXP '^[0-9]+\\.?[0-9]*$' AND l2.langtext LIKE '%alb%' AND l2.einheit LIKE 'mg/g%' AND l2.zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date) " & _
       "LEFT JOIN `labor1a` l1 ON n.pat_id = l1.pat_id AND l1.abk_ur IN ('albcre','albq','album','albup') AND l1.wert REGEXP '^[0-9]+\\.?[0-9]*$' AND l1.langtext LIKE '%alb%' AND l1.einheit LIKE 'mg/g%' AND l1.zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date) " & _
       "WHERE NOT ISNULL(n.pat_id) GROUP BY n.pat_id ORDER BY n.pat_id"
Vsql = "SELECT pat_id" & vbCrLf & _
    ",(SELECT MIN(CAST(wert AS DECIMAL(9,2))) FROM labor1a WHERE pat_id = n.pat_id AND abkü IN ('albcre','albkre','albq','album','albup') AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND langtext LIKE '%alb%' AND einheit LIKE 'mg/g%' AND zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date)) min1" & vbCrLf & _
    ",(SELECT MIN(CAST(wert AS DECIMAL(9,2))) FROM laborywert w INNER JOIN laboryus u ON w.usid = u.id  WHERE pat_id = n.pat_id AND abkü IN ('albcre','albkre','albq','album','albup') AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND langtext LIKE '%alb%' AND einheit LIKE 'mg/g%' AND eingang > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date)) min2" & vbCrLf & _
    ",(SELECT MAX(CAST(wert AS DECIMAL(9,2))) FROM labor1a WHERE pat_id = n.pat_id AND abkü IN ('albcre','albkre','albq','album','albup') AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND langtext LIKE '%alb%' AND einheit LIKE 'mg/g%' AND zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date)) max1" & vbCrLf & _
    ",(SELECT MAX(CAST(wert AS DECIMAL(9,2))) FROM laborywert w INNER JOIN laboryus u ON w.usid = u.id WHERE pat_id = n.pat_id AND abkü IN ('albcre','albkre','albq','album','albup') AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND langtext LIKE '%alb%' AND einheit LIKE 'mg/g%' AND eingang > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date)) max2" & vbCrLf & _
    "FROM namen n"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "AlbCreMinMax"
Vsql = "SELECT pat_id, IF(min1<min2 OR ISNULL(min2),min1,min2) Min, IF(max1>max2 OR ISNULL(max2),max1,max2) Max FROM _AlbCreMinMax"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

'VN = "AlbCre"
'Vsql = "SELECT n.pat_id, IF(ISNULL(l2.zeitpunkt),l1.zeitpunkt,l2.zeitpunkt) zp, IF(ISNULL(l2.zeitpunkt),l1.wert,l2.wert) wert " & _
'       "FROM namen n " & _
'       "LEFT JOIN labor2a l2 ON n.pat_id = l2.pat_id AND l2.abk_ur IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND l2.einheit LIKE 'mg/g %' " & _
'       "LEFT JOIN labor1a l1 ON n.pat_id = l1.pat_id AND l1.abk_ur IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND l1.einheit LIKE 'mg/g %' " & _
'       "WHERE (l2.Zeitpunkt = l1.Zeitpunkt OR ISNULL(l2.Zeitpunkt) OR ISNULL(l1.Zeitpunkt)) AND NOT (ISNULL(l2.Zeitpunkt) AND ISNULL(l1.Zeitpunkt)) " & _
'       "GROUP BY pat_id, l2.zeitpunkt, l1.zeitpunkt " & _
'       "ORDER BY n.pat_id, l2.zeitpunkt, l1.zeitpunkt "
'Call DtbCreateQueryDef(VN, Vsql)
'vz = vz + 1


VN = "_KreaMinMax"
'Vsql = "SELECT n.pat_id, MAX(IF(CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2))>CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)) OR ISNULL(l1.wert),CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)),CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)))) max, MIN(IF(CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)) < CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)) OR ISNULL(l1.wert),CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)), CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)))) min " & _
       "FROM `namen` n " & _
       "LEFT JOIN `labor2a` l2 ON n.pat_id = l2.pat_id AND l2.abk_ur IN ('kre02','creat','krea','krea02') AND NOT l2.langtext LIKE '%urin%' AND l2.einheit = 'mg/dl' AND l2.zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date) " & _
       "LEFT JOIN `labor1a` l1 ON n.pat_id = l1.pat_id AND l1.abk_ur IN ('kre02','creat','krea','krea02') AND NOT l1.langtext LIKE '%urin%' AND l1.einheit = 'mg/dl' AND l1.zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date) " & _
       "WHERE NOT ISNULL(n.pat_id) GROUP BY n.pat_id ORDER BY n.pat_id"
Vsql = "SELECT pat_id, geschlecht" & vbCrLf & _
    ",(SELECT MIN(CAST(wert AS DECIMAL(9,2))) FROM labor1a WHERE pat_id = n.pat_id AND abkü IN ('kre02','creat','krea','krea02') AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND NOT langtext LIKE '%urin%' AND einheit = 'mg/dl' AND zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date)) min1" & vbCrLf & _
    ",(SELECT MIN(CAST(wert AS DECIMAL(9,2))) FROM laborywert w INNER JOIN laboryus u ON w.usid = u.id WHERE pat_id = n.pat_id AND abkü IN ('kre02','creat','krea','krea02') AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND NOT langtext LIKE '%urin%' AND einheit = 'mg/dl' AND eingang > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date)) min2" & vbCrLf & _
    ",(SELECT MAX(CAST(wert AS DECIMAL(9,2))) FROM labor1a WHERE pat_id = n.pat_id AND abkü IN ('kre02','creat','krea','krea02') AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND NOT langtext LIKE '%urin%' AND einheit = 'mg/dl' AND zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date)) max1" & vbCrLf & _
    ",(SELECT MAX(CAST(wert AS DECIMAL(9,2))) FROM laborywert w INNER JOIN laboryus u ON w.usid = u.id WHERE pat_id = n.pat_id AND abkü IN ('kre02','creat','krea','krea02') AND wert REGEXP '^[0-9]+\\.?[0-9]*$' AND NOT langtext LIKE '%urin%' AND einheit = 'mg/dl' AND eingang > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date)) max2" & vbCrLf & _
    "FROM namen n"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "KreaMinMax"
Vsql = "SELECT pat_id, IF(min1<min2 OR ISNULL(min2),min1,min2) Min, IF(max1>max2 OR ISNULL(max2),max1,max2) Max FROM _KreaMinMax"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

'MDRD 4: eGFR = 186*POW(Crea,-1.154)*POW(Alter,-0.203)*IF(weibl,0.742)
'VN = "_GFRtest"
'Vsql = "SELECT n.pat_id, " & _
       "floor(pow(MAX(IF(CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)) > CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)) OR ISNULL(l1.wert),CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)),CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)))),-1.154)*186*POW(DATEdiff(IF(CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)) > CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)) OR ISNULL(l1.wert),l2.zeitpunkt,l1.zeitpunkt),n.gebdat)*0.00273792574745,-0.203)*IF(n.geschlecht='w',0.742,1)) gfr_max, " & _
       "floor(pow(min(IF(CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)) < CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)) OR ISNULL(l1.wert),CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)),CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)))),-1.154)*186*POW(DATEdiff(IF(CONVERT(REPLACE(CONCAT('0',trim(l2.wert)),',','.'),DECIMAL(9,2)) < CONVERT(REPLACE(CONCAT('0',trim(l1.wert)),',','.'),DECIMAL(9,2)) OR ISNULL(l1.wert),l2.zeitpunkt,l1.zeitpunkt),n.gebdat)*0.00273792574745,-0.203)*IF(n.geschlecht='w',0.742,1)) gfr_min " & _
       "FROM `namen` n " & _
       "LEFT JOIN `labor2a` l2 ON n.pat_id = l2.pat_id AND l2.abk_ur IN ('kre02','creat','krea','krea02') AND NOT l2.langtext LIKE '%urin%' AND l2.einheit = 'mg/dl' AND l2.zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date) " & _
       "LEFT JOIN `labor1a` l1 ON n.pat_id = l1.pat_id AND l1.abk_ur IN ('kre02','creat','krea','krea02') AND NOT l1.langtext LIKE '%urin%' AND l1.einheit = 'mg/dl' AND l1.zeitpunkt > CAST((SELECT CONCAT(YEAR(SUBDATE(NOW(),INTERVAL 200 DAY)), '-',((month( SUBDATE(NOW(),INTERVAL 200 DAY) )-1) div 3)*3 + 1,'-01')) As Date) " & _
       "WHERE NOT ISNULL(n.pat_id) GROUP BY n.pat_id ORDER BY n.pat_id"
'Call DtbCreateQueryDef(VN, Vsql)
'vz = vz + 1

VN = "_GFRMinMax"
Vsql = "SELECT pat_id, geschlecht, IF(min1<min2 OR ISNULL(min2),min1,min2) Min, IF(max1>max2 OR ISNULL(max2),max1,max2) Max FROM _KreaMinMax"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "GFRMinMax"
Vsql = "SELECT pat_id" & vbCrLf & _
       ",round(186*POW(Max,-1.154)*POW(patalter(pat_id),-0.203)*IF(geschlecht='w',0.742,1)) Min" & vbCrLf & _
       ",round(186*POW(Min,-1.154)*POW(patalter(pat_id),-0.203)*IF(geschlecht='w',0.742,1)) Max" & vbCrLf & _
       "FROM _GFRMinMax"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

'VN = "__fuerlmp"
'Vsql = "SELECT pat_id, zeitpunkt, mpnr FROM `medplan` mp GROUP BY pat_id, zeitpunkt, mpnr ORDER BY pat_id,zeitpunkt DESC, mpnr DESC"
'Call DtbCreateQueryDef(VN, Vsql)
'vz = vz + 1


'VN = "_fuerlmp"
'Vsql = "SELECT pat_id,mpnr FROM __fuerlmp GROUP BY pat_id"
'Call DtbCreateQueryDef(VN, Vsql)
'vz = vz + 1

VN = "lmp"
'Vsql = "SELECT mp.* FROM _fuerlmp i INNER JOIN `medplan` mp ON i.pat_id = mp.pat_id AND i.mpnr = mp.mpnr"
Vsql = "SELECT * FROM medplan mp " & vbCrLf & _
       "WHERE mpnr=(SELECT mpnr FROM medplan WHERE pat_id=mp.Pat_ID ORDER BY pat_id, zeitpunkt DESC,mpnr DESC LIMIT 1) " & vbCrLf & _
       "AND zeitpunkt=(SELECT zeitpunkt FROM medplan WHERE pat_id=mp.Pat_ID ORDER BY pat_id, zeitpunkt DESC,mpnr DESC LIMIT 1)"
' letzte Zeile notwendig wegen Fehler in MPNr durch Hausbesuchsimport
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "eintrhist"
Vsql = "SELECT `eintrhist1`.`ID` AS `ID`,`eintrhist1`.`Pat_ID` AS `Pat_ID`,`eintrhist1`.`ZeitPunkt` AS `ZeitPunkt`,`eintrhist1`.`Art` AS `Art`,`eintrhist1`.`Inhalt` AS `Inhalt`,`eintrhist1`.`QS` AS `QS`,`eintrhist1`.`QT` AS `QT`,`eintrhist2`.`FID` AS `FID`,`eintrhist2`.`absPos` AS `absPos`,`eintrhist2`.`AktZeit` AS `AktZeit`,`eintrhist2`.`StByte` AS `StByte` FROM (`eintrhist1` JOIN `eintrhist2` on((`eintrhist1`.`ID` = `eintrhist2`.`ID`)))"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "letze faelle"
Vsql = "SELECT `f`.`Pat_ID` AS `pat_id`,`f`.`FID` AS `fid`,`f`.`SchGr` AS `schgr`,`f`.`BhFB` AS `bhfb`,`f`.`IK` AS `ik` FROM (`_lfaelle` `l` LEFT JOIN `faelle` `f` on(((`f`.`Pat_ID` = `l`.`pid`) AND (`f`.`BhFB` = `l`.`mbhfb`))))"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

' 20.3.12: Folgende Views nach aktuellem Stand (ohne inhaltliche Prüfung) übernommen

VN = "_qsumme"
Vsql = "SELECT (COUNT(0) * IF((`e`.`Euro` = 18.75 OR e.euro = 19.05 OR euro = 14.25) AND LEFT(e.ziffer,1)='9',75,`e`.`Euro`)) AS `wert`,COUNT(0) AS `zahl`,`e`.`Euro` AS `euro`,`f`.`Quartal` AS `quartal` FROM ((`faelle` `f` LEFT JOIN `leistungen` `l` on((`f`.`FID` = `l`.`FID`))) LEFT JOIN `EBM2010` `e` on(((`l`.`Leistung` = `e`.`Ziffer`) AND (`e`.`Leistungstext` <> '')))) WHERE ((`f`.`SchGr` <> 90) AND (`l`.`Leistung` <> '00000') AND (`e`.`Euro` IS NOT NULL)) GROUP BY `l`.`Leistung`,`f`.`Quartal`"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "qsumme"
Vsql = "SELECT SUM(`_qsumme`.`wert`) `summe`,`_qsumme`.`quartal` `quartal` FROM `_qsumme` GROUP BY `_qsumme`.`quartal` ORDER BY SUBSTR(`_qsumme`.`quartal`,2) DESC,`_qsumme`.`quartal` DESC"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "CSII bei Typ 2"
Vsql = "SELECT a.Pat_ID, gesname(a.pat_id) PName, t.zp, t.therart, t.grund, d.gicd icd " & vbCrLf & _
"FROM anamnesebogen a" & vbCrLf & _
"LEFT JOIN therarten t USING (pat_id)" & vbCrLf & _
"LEFT JOIN diagview d ON d.Pat_id = a.Pat_id AND (d.gicd REGEXP '^E1[1-4]\.' OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit in ('G',' ')))" & vbCrLf & _
"WHERE (a.Ther1 = 'CSII' OR t.therart='CSII') AND d.ICD IS NOT NULL"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

' Patienten mit Diabetesneumanifestation von aktq() aus gesehen
' wenn eine Diabetes-ICD eingetragen ist und das in andm eingetragene Manifestationsjahr oder im Fall eines
' nicht eingetragenen Jahres das Jahr, in dem zum ersten Mal die Diabetesanamnese erhoben wurde,
' dem aktuellen Jahr oder im 1.Quartal dem Vorjahr entspricht, dann könnte eine Neumanifestation vorliegen
VN = "anb_neuman"
Vsql = "SELECT IF(a.DsJ='',COALESCE(erstandm,vorgestellt),a.DsJ) >= MID(aktq(),2)-IF(LEFT(aktq() COLLATE utf8_general_ci,1)='1',1,0)" & vbCrLf & _
       " AND obicd AND NOT `Diabetes seit` RLIKE 'Jahr|Jugend|Kindheit|a$|J$|l[aä]ng|weiß|^[0-9]{1,2}\.[0-9]$' obneu, a.*" & vbCrLf & _
       "FROM (SELECT CASE" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE '[ ./][0-9]{4}$' THEN RIGHT(`Diabetes seit`,4)" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE '[ -./][4-9][0-9]$' THEN CONCAT('19',RIGHT(`Diabetes seit`,2))" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE '[ -./][0-3][0-9]$' THEN CONCAT('20',RIGHT(`Diabetes seit`,2))" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE '^[0-9]{4}$' THEN `Diabetes seit`" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE '[0-9]{4}' THEN REGEXP_SUBSTR(`Diabetes seit`,'[0-9]{4}')" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE '[ -.][0-9]+ [aJ]' THEN YEAR(vorgestellt)-SUBSTRING_INDEX(MID(REGEXP_SUBSTR(`Diabetes seit`,'[ -.][0-9]+ [aJ]'),2), ' ', 1)" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE '^[0-9]+ [aJ]' THEN YEAR(vorgestellt)-SUBSTRING_INDEX(REGEXP_SUBSTR(`Diabetes seit`,'^[0-9]+ [aJ]'), ' ', 1)" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE '^[0-9]+a' THEN YEAR(vorgestellt)-SUBSTRING_INDEX(MID(REGEXP_SUBSTR(`Diabetes seit`,'^[0-9]+a'),2), 'a', 1)" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE '[ -.][0-9]+ Mon' THEN YEAR(vorgestellt-INTERVAL SUBSTRING_INDEX(MID(REGEXP_SUBSTR(`Diabetes seit`,'[ -.][0-9]+ Mon'),2), ' ', 1) MONTH)" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE '^[0-9]+ Mon' THEN YEAR(vorgestellt-INTERVAL SUBSTRING_INDEX(REGEXP_SUBSTR(`Diabetes seit`,'^[0-9]+ Mon'), ' ', 1) MONTH)" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE '[0-9]/[4-9][0-9]' THEN CONCAT('19',MID(REGEXP_SUBSTR(`Diabetes seit`,'[0-9]/[0-9]{2}'),3))" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE '[0-9]/[0-3][0-9]' THEN CONCAT('20',MID(REGEXP_SUBSTR(`Diabetes seit`,'[0-9]/[0-9]{2}'),3))" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE 'heute|^bu$|tag|^[-n]$|noch nicht|neu' THEN YEAR(vorgestellt)" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE 'woche' THEN YEAR(vorgestellt)-IF(MONTH(vorgestellt)=1,1,0)" & vbCrLf & _
       "WHEN `Diabetes seit` RLIKE '^[0-9]{1,2}$' THEN YEAR(vorgestellt)-`Diabetes seit`" & vbCrLf & _
       "ELSE '' END DsJ,`Diabetes seit`,pid IS NOT NULL obicd,a.pat_id,vorgestellt,bmi" & vbCrLf & _
       ",(SELECT SUBSTRING_INDEX(GROUP_CONCAT(zeitpunkt ORDER BY zeitpunkt),',',1) FROM eintraege WHERE art LIKE 'andm%' AND pat_id=a.pat_id) erstandm" & vbCrLf & _
       " FROM anamnesebogen a" & vbCrLf & _
       " LEFT JOIN (SELECT pat_id pid FROM diagview WHERE gicd LIKE 'E1%' GROUP BY pat_id) d ON d.pid=a.pat_id" & vbCrLf & _
       ") a;"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


VN = "fallzahlstand_1"
Vsql = "SELECT COUNT(0) AS `zahl`,(COUNT(0) - COUNT(`f`.`erst`)) AS `z0`,`f`.`Quartal` AS `quartal` FROM `faelleverschiedenneu` `f` WHERE ((`f`.`SchGr` <> '90') AND ((to_days(`f`.`Fanf`) - to_days(CONCAT(SUBSTR(`f`.`Quartal`,2,4),'-',(((LEFT(`f`.`Quartal`,1) - 1) * 3) + 1),'-01'))) BETWEEN 0 AND (to_days((NOW() - INTERVAL 0 DAY)) - to_days(CONCAT(YEAR((NOW() - INTERVAL 0 DAY)),'-',((((month((NOW() - INTERVAL 0 DAY)) - 1) DIV 3) * 3) + 1),'-01')))) AND ((`f`.`Pat_ID` < 3044) OR (`f`.`Pat_ID` > 50000))) GROUP BY `f`.`Quartal` ORDER BY SUBSTR(`f`.`Quartal`,2,4),LEFT(`f`.`Quartal`,1)"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "fallzahlstand_2"
Vsql = "SELECT COUNT(0) AS `zahl`,(COUNT(0) - COUNT(`f`.`erst`)) AS `z0`,`f`.`Quartal` AS `quartal` FROM `faelleverschiedenneu` `f` WHERE ((`f`.`SchGr` <> '90') AND ((to_days(`f`.`Fanf`) - to_days(CONCAT(SUBSTR(`f`.`Quartal`,2,4),'-',(((LEFT(`f`.`Quartal`,1) - 1) * 3) + 1),'-01'))) BETWEEN 0 AND (to_days((NOW() - INTERVAL 0 DAY)) - to_days(CONCAT(YEAR((NOW() - INTERVAL 0 DAY)),'-',((((month((NOW() - INTERVAL 0 DAY)) - 1) DIV 3) * 3) + 1),'-01')))) AND (`f`.`Pat_ID` > 3044)) GROUP BY `f`.`Quartal` ORDER BY SUBSTR(`f`.`Quartal`,2,4),LEFT(`f`.`Quartal`,1)"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "fallzahlstand_3"
Vsql = "SELECT COUNT(0) AS `zahl`,(COUNT(0) - COUNT(`f`.`erst`)) AS `z0`,`f`.`Quartal` AS `quartal` FROM `faelleverschiedenneu` `f` WHERE ((`f`.`SchGr` <> '90') AND ((to_days(`f`.`Fanf`) - to_days(CONCAT(SUBSTR(`f`.`Quartal`,2,4),'-',(((LEFT(`f`.`Quartal`,1) - 1) * 3) + 1),'-01'))) BETWEEN 0 AND (to_days((NOW() - INTERVAL 0 DAY)) - to_days(CONCAT(YEAR((NOW() - INTERVAL 0 DAY)),'-',((((month((NOW() - INTERVAL 0 DAY)) - 1) DIV 3) * 3) + 1),'-01')))) AND 1 AND (SUBSTR(`f`.`Quartal`,2,4) > '2008') AND (`f`.`Quartal` <> '12009')) GROUP BY `f`.`Quartal` ORDER BY SUBSTR(`f`.`Quartal`,2,4),LEFT(`f`.`Quartal`,1)"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "kontoausw"
Vsql = "SELECT `e`.`eingid` AS `eingid`,`e`.`tabelle` AS `tabelle`,CAST(`e`.`fdt` As Date) AS `datum`,`e`.`DATEi` AS `DATEi`,IF((LEFT(`e`.`DATEi`,1) BETWEEN '0' AND '9'),LEFT(`e`.`DATEi`,(locate('_',`e`.`DATEi`) - 1)),IF((locate('_',`e`.`DATEi`,(locate('_',`e`.`DATEi`) + 1)) <> 0),SUBSTR(`e`.`DATEi`,(locate('_',`e`.`DATEi`) + 1),((locate('_',`e`.`DATEi`,(locate('_',`e`.`DATEi`) + 1)) - LOCATE('_',`e`.`DATEi`)) - 1)),IF((locate(' ',`e`.`DATEi`,(locate(' ',`e`.`DATEi`) + 1)) <> 0),SUBSTR(`e`.`DATEi`,(locate(' ',`e`.`DATEi`) + 1),((locate(' ',`e`.`DATEi`,(locate(' ',`e`.`DATEi`) + 1)) - LOCATE(' ',`e`.`DATEi`)) - 1)),'?'))) AS `Konto` FROM `eingelesen` `e`"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

VN = "Schulungen akt. Quartal"
Vsql = "SELECT `e`.`Pat_ID` AS `pat_id`,CONCAT(`n`.`Nachname`,' ',`n`.`Vorname`) AS `name`,`e`.`ZeitPunkt` AS `zeitpunkt`,`e`.`Inhalt` AS `inhalt`,`e`.`QS` AS `qs` FROM (`eintraege` `e` LEFT JOIN `namen` `n` on((`e`.`Pat_ID` = `n`.`Pat_ID`))) WHERE ((`e`.`Art` = 'schul') AND (`e`.`ZeitPunkt` " & DiesQ() & ")) ORDER BY `e`.`ZeitPunkt`"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

'VN = "test0"
'Vsql = "SELECT `u`.`Pat_id` AS `Pat_ID`,CAST(`u`.`Eingang` As Date) AS `zeitpunkt`,`u`.`BefArt` AS `FertigStGrad`,IF(ISNULL(`n2`.`Abkü`),`w`.`Abkü`,`n2`.`Abkü`) AS `Abkü`,`w`.`Abkü` AS `abk_ur`,`w`.`Langtext` AS `Langtext`,`w`.`Wert` AS `Wert`,IF(ISNULL(`n2`.`Abkü`),`w`.`Einheit`,`n2`.`Einheit`) AS `Einheit`,`w`.`Einheit` AS `Einheit_ur`,CONCAT(IF((erkl.text REGEXP '^:[ /*:]*$'),'',IF((erkl.text REGEXP '^:[ /*]*:'),CONCAT(SUBSTR(erkl.text,(locate(':',erkl.text,2) + 1)),';'),IF((erkl.text = '.'),'',IF((erkl.text = ''),'',CONCAT(erkl.text,';'))))),`w`.`Kommentar`) AS `Kommentar`,IF(ISNULL(`n2`.`Abkü`),`nb`.`NB`,`nb2`.`NB`) AS `NB`,`nb`.`NB` AS `NB_ur`,IF(ISNULL(`n2`.`Abkü`),`nb`.`uNg`,`nb2`.`uNg`) AS `uNg`,`nb`.`uNg` AS `uNg_ur`,IF(ISNULL(`n2`.`Abkü`),`nb`.`oNg`,`nb2`.`oNg`) AS `oNg`,`nb`.`oNg` AS `oNg_ur`,`s`.`Labor` AS `Labor` FROM `laboryus` `u` LEFT JOIN `namen` `na` ON `u`.`Pat_id` = `na`.`Pat_ID` LEFT JOIN `laborywert` `w` " & _
'"on `u`.`id` = `w`.`usid` LEFT JOIN `laborypnb` `nb` ON `w`.`nbid` = `nb`.`id` LEFT JOIN `laborypneu` `neu` ON `nb`.`pid` = `neu`.`id` LEFT JOIN `laborypgl` `gl` ON `neu`.`id` = `gl`.`idxpneu` LEFT JOIN `laborypneu` `n2` ON `gl`.`idxpbez` = `n2`.`id` LEFT JOIN `laborypnb` `nb2` ON `n2`.`id` = `nb2`.`pid` AND `nb2`.`Geschlecht` IN (3,9,0,IF((`na`.`Geschlecht` = 'm'),IF(((`na`.`GebDat` + INTERVAL 18 YEAR) > CAST(`u`.`Eingang` As Date)),1,4),IF(((`na`.`GebDat` + INTERVAL 18 YEAR) > CAST(`u`.`Eingang` As Date)),2,5))) LEFT JOIN `laborysaetze` `s` ON `u`.`SatzID` = `s`.`SatzID`  LEFT JOIN laboryhinw erkl ON erkl.id=w.erklid WHERE (`u`.`Pat_id` = 14)"
'Call DtbCreateQueryDef(VN, Vsql)
'vz = vz + 1

'VN = "test2"
'Vsql = "SELECT `nb2`.`Eingang` AS `eingang`,`u`.`Pat_id` AS `Pat_ID`,CAST(`u`.`Eingang` As Date) AS `zeitpunkt`,`u`.`BefArt` AS `FertigStGrad`,IF(ISNULL(`n2`.`Abkü`),`w`.`Abkü`,`n2`.`Abkü`) AS `Abkü`,`w`.`Abkü` AS `abk_ur`,`w`.`Langtext` AS `Langtext`,`w`.`Wert` AS `Wert`,IF(ISNULL(`n2`.`Abkü`),`w`.`Einheit`,`n2`.`Einheit`) AS `Einheit`,`w`.`Einheit` AS `Einheit_ur`,CONCAT(IF((erkl.text REGEXP '^:[ /*:]*$'),'',IF((erkl.text REGEXP '^:[ /*]*:'),CONCAT(SUBSTR(erkl.text,(locate(':',erkl.text,2) + 1)),';'),IF((erkl.text = '.'),'',IF((erkl.text = ''),'',CONCAT(erkl.text,';'))))),`w`.`Kommentar`) AS `Kommentar`,IF(ISNULL(`n2`.`Abkü`),`nb`.`NB`,`nb2`.`NB`) AS `NB`,`nb`.`NB` AS `NB_ur`,IF(ISNULL(`n2`.`Abkü`),`nb`.`uNg`,`nb2`.`uNg`) AS `uNg`,`nb`.`uNg` AS `uNg_ur`,IF(ISNULL(`n2`.`Abkü`),`nb`.`oNg`,`nb2`.`oNg`) AS `oNg`,`nb`.`oNg` AS `oNg_ur`,`s`.`Labor` AS `Labor` FROM ((((((((`laboryus` `u` LEFT JOIN `namen` `na` on((`u`.`Pat_id` = `na`.`Pat_ID`))) " & _
'"LEFT JOIN `laborywert` `w` on((`u`.`id` = `w`.`usid`))) LEFT JOIN `laborypnb` `nb` on((`w`.`nbid` = `nb`.`id`))) LEFT JOIN `laborypneu` `neu` on((`nb`.`pid` = `neu`.`id`))) LEFT JOIN `laborypgl` `gl` on((`neu`.`id` = `gl`.`idxpneu`))) LEFT JOIN `laborypneu` `n2` on((`gl`.`idxpbez` = `n2`.`id`))) LEFT JOIN `laborypnb` `nb2` on(((`n2`.`id` = `nb2`.`pid`) AND (`nb2`.`Geschlecht` IN (3,9,0,IF((`na`.`Geschlecht` = 'm'),IF(((`na`.`GebDat` + INTERVAL 18 YEAR) > CAST(`u`.`Eingang` As Date)),1,4),IF(((`na`.`GebDat` + INTERVAL 18 YEAR) > CAST(`u`.`Eingang` As Date)),2,5))))))) LEFT JOIN `laborysaetze` `s` on((`u`.`SatzID` = `s`.`SatzID`))) WHERE ((`u`.`Pat_id` = 14) AND (`n2`.`Abkü` = 'alkp02') AND (`nb2`.`Eingang` = (SELECT MAX(`laborypnb`.`Eingang`) FROM `laborypnb` WHERE (`laborypnb`.`pid` = `n2`.`id`))))  LEFT JOIN laboryhinw erkl ON erkl.id=w.erklid GROUP BY `w`.`usid`,`w`.`Abkü` ORDER BY `nb2`.`Eingang` DESC"
'Call DtbCreateQueryDef(VN, Vsql)
'vz = vz + 1

'VN = "test2"
'Vsql = "SELECT `nb2`.`Eingang` AS `eingang`,`u`.`Pat_id` AS `Pat_ID`,CAST(`u`.`Eingang` As Date) AS `zeitpunkt`,`u`.`BefArt` AS `FertigStGrad`,IF(ISNULL(`n2`.`Abkü`),`w`.`Abkü`,`n2`.`Abkü`) AS `Abkü`,`w`.`Abkü` AS `abk_ur`,`w`.`Langtext` AS `Langtext`,`w`.`Wert` AS `Wert`,IF(ISNULL(`n2`.`Abkü`),`w`.`Einheit`,`n2`.`Einheit`) AS `Einheit`,`w`.`Einheit` AS `Einheit_ur`,CONCAT(IF((`w`.`Erklärung` REGEXP '^:[ /*:]*$'),'',IF((`w`.`Erklärung` REGEXP '^:[ /*]*:'),CONCAT(SUBSTR(`w`.`Erklärung`,(locate(':',`w`.`Erklärung`,2) + 1)),';'),IF((`w`.`Erklärung` = '.'),'',IF((`w`.`Erklärung` = ''),'',CONCAT(`w`.`Erklärung`,';'))))),`w`.`Kommentar`) AS `Kommentar`,IF(ISNULL(`n2`.`Abkü`),`nb`.`NB`,`nb2`.`NB`) AS `NB`,`nb`.`NB` AS `NB_ur`,IF(ISNULL(`n2`.`Abkü`),`nb`.`uNg`,`nb2`.`uNg`) AS `uNg`,`nb`.`uNg` AS `uNg_ur`,IF(ISNULL(`n2`.`Abkü`),`nb`.`oNg`,`nb2`.`oNg`) AS `oNg`,`nb`.`oNg` AS `oNg_ur`,`s`.`Labor` AS `Labor` FROM ((((((((`laboryus` `u` LEFT JOIN `namen` `na` on((`u`.`Pat_id` = `na`.`Pat_ID`))) " & _
'"LEFT JOIN `laborywert` `w` on((`u`.`id` = `w`.`usid`))) LEFT JOIN `laborypnb` `nb` on((`w`.`nbid` = `nb`.`id`))) LEFT JOIN `laborypneu` `neu` on((`nb`.`pid` = `neu`.`id`))) LEFT JOIN `laborypgl` `gl` on((`neu`.`id` = `gl`.`idxpneu`))) LEFT JOIN `laborypneu` `n2` on((`gl`.`idxpbez` = `n2`.`id`))) LEFT JOIN `laborypnb` `nb2` on(((`n2`.`id` = `nb2`.`pid`) AND (`nb2`.`Geschlecht` IN (3,9,0,IF((`na`.`Geschlecht` = 'm'),IF(((`na`.`GebDat` + INTERVAL 18 YEAR) > CAST(`u`.`Eingang` As Date)),1,4),IF(((`na`.`GebDat` + INTERVAL 18 YEAR) > CAST(`u`.`Eingang` As Date)),2,5))))))) LEFT JOIN `laborysaetze` `s` on((`u`.`SatzID` = `s`.`SatzID`))) WHERE ((`u`.`Pat_id` = 14) AND (`n2`.`Abkü` = 'alkp02') AND (`nb2`.`Eingang` = (SELECT MAX(`laborypnb`.`Eingang`) FROM `laborypnb` WHERE (`laborypnb`.`pid` = `n2`.`id`)))) GROUP BY `w`.`usid`,`w`.`Abkü` ORDER BY `nb2`.`Eingang` DESC"
'Call DtbCreateQueryDef(VN, Vsql)
'vz = vz + 1

VN = "versorgungsamt oberfranken"
Vsql = "SELECT `i`.`titel` AS `titel`,`i`.`tsid` AS `tsid`,`i`.`transe` AS `transe`,`i`.`transs` AS `transs`,`i`.`id` AS `id`,`i`.`fsize` AS `fsize`,`i`.`pages` AS `pages`,`i`.`devname` AS `devname`,`i`.`retries` AS `retries`,`i`.`csid` AS `csid`,`i`.`routi` AS `routi`,`i`.`callerid` AS `callerid` FROM `inca` `i` WHERE ((`i`.`tsid` LIKE '%803599%') AND (`i`.`pages` > 1)) ORDER BY `i`.`transe` DESC"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1

' wer durch Ausprobieren ermittelt am 1.11.22 mit SELECT * FROM anaktk WHERE wer=''
VN = "anaktk"
Vsql = "" & _
"SELECT CASE " & vbCrLf & _
"WHEN (obtk OR (NOT obtk AND NOT obgs AND NOT obdw AND NOT obah)) AND (COALESCE(tkbis>gsbis-INTERVAL 3 MONTH,1) OR gszahl<5) AND COALESCE(tkbis>ahbis-INTERVAL 3 MONTH,1) THEN 'tk'" & vbCrLf & _
"WHEN (obgs OR (NOT obtk AND NOT obgs AND NOT obdw AND NOT obah)) AND COALESCE(gsbis>tkbis-INTERVAL 3 MONTH,1) AND COALESCE(gsbis>ahbis-INTERVAL 3 MONTH,1) THEN 'gs'" & vbCrLf & _
"WHEN (obdw AND NOT obtk AND NOT obgs AND NOT obah) AND tkzahl<6 AND gszahl<6 AND ahzahl=0 THEN '-'" & vbCrLf & _
"WHEN (obah OR (NOT obtk AND NOT obgs AND NOT obdw AND NOT obah)) AND COALESCE(ahbis>tkbis-INTERVAL 3 MONTH,1) AND COALESCE(ahbis>gsbis-INTERVAL 3 MONTH,1) THEN 'ah'" & vbCrLf & _
"WHEN tkbis>gsbis AND tkbis>ahbis AND tkzahl>0 THEN 'tk'" & vbCrLf & _
"WHEN gsbis>tkbis AND gsbis>ahbis AND gszahl>5 THEN 'gs'" & vbCrLf & _
"WHEN ahbis>tkbis AND ahbis>tkbis AND ahzahl>0 THEN 'ah'" & vbCrLf & _
"WHEN obtk AND NOT obgs AND NOT obdw AND NOT obah AND gszahl<6 AND dwzahl=0 AND ahzahl=0 THEN 'tk'" & vbCrLf & _
"WHEN obah AND NOT obgs AND NOT obtk AND NOT obdw AND gszahl<6 AND dwzahl=0 AND ahzahl=0 THEN 'ah'" & vbCrLf & _
"" & vbCrLf & _
"" & vbCrLf & _
"WHEN NOT obtk AND NOT obgs AND NOT obdw AND NOT obah AND tkbis=0 AND gsbis=0 AND dwbis=0 AND ahbis=0 THEN '-'" & vbCrLf & _
"ELSE '' END wer" & vbCrLf
Vsql = Vsql & _
", i.* FROM (SELECT a.`Prim` `Prim`,a.`Pat_id` PID" & vbCrLf & _
", COALESCE((SELECT '1' FROM desktop WHERE pat_id = a.pat_id AND iconpath RLIKE '4eckblau' AND showasnote=0 LIMIT 1),'') obtk " & vbCrLf & _
", COALESCE((SELECT '1' FROM desktop WHERE pat_id = a.pat_id AND iconpath RLIKE '4eckgelb' AND showasnote=0 LIMIT 1),'') obgs " & vbCrLf & _
", COALESCE((SELECT '1' FROM desktop WHERE pat_id = a.pat_id AND iconpath RLIKE '4eckmagneta' AND showasnote=0 LIMIT 1),'') obdw " & vbCrLf & _
", COALESCE((SELECT '1' FROM desktop WHERE pat_id = a.pat_id AND iconpath RLIKE '4EckHellgruen' AND showasnote=0 LIMIT 1),'') obah " & vbCrLf & _
", COALESCE((SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id=a.pat_id AND art='tk'),0) tkbis" & vbCrLf & _
", COALESCE((SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id=a.pat_id AND art='gs'),0) gsbis" & vbCrLf & _
", COALESCE((SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id=a.pat_id AND art='dw'),0) dwbis" & vbCrLf & _
", COALESCE((SELECT MAX(zeitpunkt) FROM eintraege WHERE pat_id=a.pat_id AND art='ah'),0) ahbis" & vbCrLf & _
", (SELECT COUNT(0) FROM eintraege WHERE pat_id=a.pat_id AND art='tk') tkzahl" & vbCrLf & _
", (SELECT COUNT(0) FROM eintraege WHERE pat_id=a.pat_id AND art='gs') gszahl" & vbCrLf & _
", (SELECT COUNT(0) FROM eintraege WHERE pat_id=a.pat_id AND art='dw') dwzahl" & vbCrLf & _
", (SELECT COUNT(0) FROM eintraege WHERE pat_id=a.pat_id AND art='ah') ahzahl" & vbCrLf & _
"FROM anamnesebogen a" & vbCrLf & _
"GROUP BY a.pat_id) i"


'", (SELECT COUNT(0) FROM eintraege WHERE pat_id=a.pat_id AND (art='tk' OR (inhalt LIKE '%(tk)%'))) tkzahl" & vbCrLf & _
", (SELECT COUNT(0) FROM eintraege WHERE pat_id=a.pat_id AND (art='gs' OR (inhalt LIKE '%(gs)%'))) gszahl" & vbCrLf & _

' ", (SELECT COUNT(0) FROM eintraege WHERE pat_id=a.pat_id AND (art='tk' OR (art IN ('angd','aufgd','auto','ba','bla','bzvgl','colo','debr','doppler','duplex','ictauf','ih','ik','inj','insgd','kv','ni','rrvgl','sono','tst','tv','vac','vkgd','vkgd2','wr','wv') AND inhalt LIKE '%(tk)%'))) tkzahl" & vbCrLf & _
", (SELECT COUNT(0) FROM eintraege WHERE pat_id=a.pat_id AND (art='gs' OR (art IN ('angd','aufgd','auto','ba','bla','bzvgl','colo','debr','doppler','duplex','ictauf','ih','ik','inj','insgd','kv','ni','rrvgl','sono','tst','tv','vac','vkgd','vkgd2','wr','wv') AND inhalt LIKE '%(gs)%'))) gszahl" & vbCrLf & _


Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1
' ", COALESCE((SELECT voret FROM sws WHERE pat_id=a.pat_id AND voret>qanf() ORDER BY voret DESC LIMIT 1),'') voret " & vbCrLf & _
' "FROM (`anamnesebogen` a LEFT JOIN `usdm` u ON (((a.`Pat_id` = u.`Pat_ID`) AND (u.`ZeitPunkt` = (SELECT MAX(`usdm`.`ZeitPunkt`) FROM `usdm` WHERE (`usdm`.`Pat_ID` = a.`Pat_id`))))));"

VN = "anakt"
Vsql = "" & _
"SELECT a.`Prim` `Prim`,a.`Pat_id` `Pat_id`,a.`Nachname` `Nachname`,a.`Vorname` `Vorname`,a.`NVorsatz` `NVorsatz`,a.`Titel` `Titel`,a.`Anrede` `Anrede`,a.`GebDat` `GebDat`,a.`Tkz` `Tkz`,a.`Versicherungsart` `Versicherungsart`,a.`Diabetestyp` `Diabetestyp`,a.`Diabetes seit` `Diabetes seit`,a.`Tabletten seit` `Tabletten seit`,a.`Insulin seit` `Insulin seit`,a.`Grund für Vorstellung` `Grund für Vorstellung`,a.`Familienanamnese` `Familienanamnese`,a.`Größe` `Größe`,a.`Gewicht` `Gewicht`,a.`bmi` `bmi`,a.`Tendenz` `Tendenz`,a.`DiabetesMedikament 1` `DiabetesMedikament 1`,a.`DiabetesMedikament 1 Menge` `DiabetesMedikament 1 Menge`,a.`DiabetesMedikament 2` `DiabetesMedikament 2`,a.`DiabetesMedikament 2 Menge` `DiabetesMedikament 2 Menge`,a.`DiabetesMedikament 3` `DiabetesMedikament 3`,a.`DiabetesMedikament 3 Menge` `DiabetesMedikament 3 Menge`,a.`DiabetesMedikament 4` " & vbCrLf & _
"`DiabetesMedikament 4`,a.`DiabetesMedikament 4 Menge` `DiabetesMedikament 4 Menge`,a.`Insulinpumpe` `Insulinpumpe`,a.`Insulinpumpe seit` `Insulinpumpe seit`,a.`Insulinpumpe Marke` `Insulinpumpe Marke`,a.`Broteinheiten gesamt` `Broteinheiten gesamt`,a.`Broteinheiten früh` `Broteinheiten früh`,a.`Broteinheiten ZM früh` `Broteinheiten ZM früh`,a.`Broteinheiten mittags` `Broteinheiten mittags`,a.`Broteinheiten nachmittags` `Broteinheiten nachmittags`,a.`Broteinheiten abends` `Broteinheiten abends`,a.`Broteinheiten nachts` `Broteinheiten nachts`,a.`Essenszeit früh` `Essenszeit früh`,a.`Essenszeit vormittags` `Essenszeit vormittags`,a.`Essenszeit mittags` `Essenszeit mittags`,a.`Essenszeit nachmittags` `Essenszeit nachmittags`,a.`Essenszeit abends` `Essenszeit abends`,a.`Essenszeit spät` `Essenszeit spät`,a.`Spritz-Eß-AbstAND früh` `Spritz-Eß-AbstAND früh`,a.`Spritz-Eß-AbstAND mittags` `Spritz-Eß-AbstAND mittags`, " & vbCrLf & _
"a.`Spritz-Eß-AbstAND abends` `Spritz-Eß-AbstAND abends`,a.`Spritzstelle früh` `Spritzstelle früh`,a.`Spritzstelle mittags` `Spritzstelle mittags`,a.`Spritzstelle abends` `Spritzstelle abends`,a.`Spritzstelle nachts` `Spritzstelle nachts`,a.`Jahr letzte Diabetesschulung` `Jahr letzte Diabetesschulung`,a.`Ort Schulung` `Ort Schulung`,a.`letztes HbA1c` `letztes HbA1c`,a.`gemessen am` `gemessen am`,a.`vorherige Werte` `vorherige Werte`,a.`BZMessungen selbst` `BZMessungen selbst`,a.`Gerät` `Gerät`,a.`BZMessungen pW` `BZMessungen pW`,a.`BZMessungen pW ndE` `BZMessungen pW ndE`,a.`BZMessungen p W nachts` `BZMessungen p W nachts`,a.`Aufschreiben` `Aufschreiben`,a.`BZWerte v d Essen` `BZWerte v d Essen`,a.`BZWerte n d Essen` `BZWerte n d Essen`,a.`UZ Tageszeit` `UZ Tageszeit`,a.`Unterzucker pM` `Unterzucker pM`,a.`UZ rechtzeitig` `UZ rechtzeitig`, " & vbCrLf & _
"a.`Fremde Hilfe pa` `Fremde Hilfe pa`,a.`Bewußtlos pa` `Bewußtlos pa`,a.`Keto pa` `Keto pa`,a.`BZgr300 pM` `BZgr300 pM`," & vbCrLf & _
"a.`Bluthochdruck` `Bluthochdruck`,a.`BHD seit` `BHD seit`,a.`BHD beh mit` `BHD beh mit`,a.`Blutdruckwerte` `Blutdruckwerte`,a.`BDselbst` `BDselbst`,a.`Schwanger` `Schwanger`,a.`Schwanger seit` `Schwanger seit`,a.`Augensp zuletzt` `Augensp zuletzt`,a.`Augensp Befund` `Augensp Befund`,a.`Netzhaut gelasert` `Netzhaut gelasert`,a.`Sehminderung unbehebbar` `Sehminderung unbehebbar`,a.`Diabet Nierenschaden` `Diabet Nierenschaden`,a.`Albumin zuletzt` `Albumin zuletzt`,a.`erhöht?` `erhöht?`,a.`Dialyse` `Dialyse`,a.`Dialyse seit` `Dialyse seit`,a.`andere Nierenerkrankung` `andere Nierenerkrankung`,a.`Herzkrankheit` `Herzkrankheit`,a.`Angina pectoris` `Angina pectoris`,a.`Herzinfarkt` `Herzinfarkt`,a.`Herzinfarkt wann` `Herzinfarkt wann`,a.`PTCA oder Stent` `PTCA oder Stent`,a.`Bypass kardial` `Bypass kardial`,a.`Bypass wann` `Bypass wann`,a.`Herzschwäche` `Herzschwäche`," & vbCrLf & _
"a.`Herzkrankheit Beschreibung` `Herzkrankheit Beschreibung`,a.`Hirndurchblutungsstörung` `Hirndurchblutungsstörung`,a.`Schlaganfall` `Schlaganfall`,a.`Beindurchblutungsstörung` `Beindurchblutungsstörung`,a.`Schaufensterkrankheit` `Schaufensterkrankheit`,a.`Bypaß peripher` `Bypaß peripher`,a.`Geschwür` `Geschwür`,a.`Amputation` `Amputation`,a.`pAVK Beschreibung` `pAVK Beschreibung`,a.`Ameisenlaufen` `Ameisenlaufen`,a.`Ameisen Ausmaß` `Ameisen Ausmaß`,a.`Druckstellen` `Druckstellen`,a.`Verformungen` `Verformungen`,a.`Verformungen Beschreibung` `Verformungen Beschreibung`,a.`Fußpflege` `Fußpflege`,a.`Podologie` `Podologie`,a.`Einlagen` `Einlagen`,a.`Neue Fußkomplikationen` `Neue Fußkomplikationen`,a.`Entleerungsstörungen Magen` `Entleerungsstörungen Magen`,a.`Entleerungsstörungen Harnblase` `Entleerungsstörungen Harnblase`,a.`Schwindel Aufstehen` `Schwindel Aufstehen`," & vbCrLf & _
"a.`Folgeerkrankungen Haut` `Folgeerkrankungen Haut`,a.`Bewegungseinschränkungen` `Bewegungseinschränkungen`,a.`Sexualstörung` `Sexualstörung`,a.`Sexualstörung seit` `Sexualstörung seit`,a.`Weitere Anamnese` `Weitere Anamnese`,a.`Tabak` `Tabak`,a.`tabakex` `tabakex`,a.`tabakbis` `tabakbis`,a.`tabakakt` `tabakakt`,a.`tabakmenge` `tabakmenge`,a.`Alkohol` `Alkohol`,a.`Mitarbeiter` `Mitarbeiter`,a.`Weitere Medikation` `Weitere Medikation`,u.`Spritzst` `Liphypertrophien Abdomen`,a.`Liphypertrophien Beine` `Liphypertrophien Beine`,a.`Liphypertrophien Arme` `Liphypertrophien Arme`,CONCAT(u.`Fußbef_re`,' | ',u.`Fußbef_li`) `Beinbefund`,CONCAT(u.`Hyperk_re`,' | ',u.`Hyperk_li`) `Hyperkeratosen`,CONCAT(u.`Ulcera_re`,' | ',u.`Ulcera_li`)  `Ulcera`,CONCAT(u.`Kraft_Zh_re`,' | ',u.`Kraft_Zh_li`) `Kraft Zehenheber`," & vbCrLf & _
"CONCAT(u.`Kraft_Zb_re`,' | ',u.`Kraft_Zb_li`)  `Kraft Zehenbeuger`, CONCAT(u.`Kraft_Knie_re`,' | ',u.`Kraft_Knie_li`)  `Kraft Knie`," & vbCrLf & _
"CONCAT(u.`ASR_re`,' | ',u.`ASR_li`) `ASR`,CONCAT(u.`PSR_re`,' | ',u.`PSR_li`)  `PSR`,CONCAT(u.`Oberfl_re`,' | ',u.`Oberfl_li`)  `Oberflächensensibilität`,CONCAT(u.`MF_re`,' | ',u.`MF_li`)  `Monofilamenttest`,CONCAT(u.`KW_re`,' | ',u.`KW_li`)  `Kalt-Warm`,CONCAT(u.`Vibr_IK_re`,' | ',u.`Vibr_IK_li`)  `Vibration IK`,CONCAT(u.`Vibr_GZ_re`,' | ',u.`Vibr_GZ_li`)  `Vibration Großzehe`,CONCAT(u.`PulsL_re`,' | ',u.`PulsL_li`)  `Puls Leiste`,CONCAT(u.`PulsKK_re`,' | ',u.`PulsKK_li`) `Puls Kniekehle`,CONCAT(u.`PulsAtp_re`,' | ',u.`PulsAtp_li`) `Puls Atp`,CONCAT(u.`PulsAdp_re`,' | ',u.`PulsAdp_li`)  `Puls Adp`,a.`RR` `RR`,a.`RRTurboMed` `RRTurboMed`,a.`Herz` `Herz`,a.`Lunge` `Lunge`,a.`Bauch` `Bauch`,a.`WS` `WS`,a.`NL` `NL`,a.`SD` `SD`,a.`Carotiden` `Carotiden`,a.`NNH` `NNH`,a.`Zähne` `Zähne`,a.`Mundhöhle` `Mundhöhle`," & vbCrLf & _
"a.`LK` `LK`,a.`BeinödVen` `BeinödVen`,a.`Neuro sonst` `Neuro sonst`,a.`Weitere Befunde` `Weitere Befunde`,a.`Schulung` `Schulung`,a.`DMP` `DMP`,a.`DMSchulz` `DMSchulz`,a.`DMSchL` `DMSchL`,a.`RRSchulz` `RRSchulz`,a.`DMPhier` `DMPhier`,a.`HANr` `HANr`,a.`HANr2` `HANr2`,a.`letzte Änderung` `letzte Änderung`,a.`Diagnosen` `Diagnosen`,a.`Vorgestellt` `Vorgestellt`,a.`Versicherung` `Versicherung`,a.`AktZeit` `AktZeit`,a.`Ther1` `Ther1`,a.`TherAkt` `TherAkt`,a.`obAn1eing` `obAn1eing`,a.`obAn2eing` `obAn2eing`,a.`obAnAeing` `obAnAeing`,a.`obCheck` `obCheck`,a.`obBZausgew` `obBZausgew`,a.`obOSaufgek` `obOSaufgek`,a.`obPodAufgek` `obPodAufgek`,a.`obMBlAusgeh` `obMBlAusgeh`,a.`obSchulaufgek` `obSchulaufgek`,a.`obDMPaufgekl` `obDMPaufgekl`,a.`obMedNetz` `obMedNetz`,a.`Hausarzt` `Hausarzt`,a.`ob` `ob`,a.`QS` `QS`,a.`QT` `QT` " & vbCrLf & _
", COALESCE((SELECT voret FROM sws WHERE pat_id=a.pat_id AND voret>qanf() ORDER BY voret DESC LIMIT 1),'') voret " & vbCrLf & _
", COALESCE((SELECT 1 FROM desktop WHERE pat_id = a.pat_id AND iconpath RLIKE '4eckblau' AND showasnote=0 LIMIT 1),0) obtk " & vbCrLf & _
", COALESCE((SELECT 1 FROM desktop WHERE pat_id = a.pat_id AND iconpath RLIKE '4eckgelb' AND showasnote=0 LIMIT 1),0) obgs " & vbCrLf & _
", COALESCE((SELECT 1 FROM desktop WHERE pat_id = a.pat_id AND iconpath RLIKE '4eckmagneta' AND showasnote=0 LIMIT 1),0) obdw " & vbCrLf & _
", COALESCE((SELECT 1 FROM desktop WHERE pat_id = a.pat_id AND iconpath RLIKE '4EckHellgruen' AND showasnote=0 LIMIT 1),0) obah " & vbCrLf & _
"FROM (`anamnesebogen` a LEFT JOIN `usdm` u ON (((a.`Pat_id` = u.`Pat_ID`) AND (u.`ZeitPunkt` = (SELECT MAX(`usdm`.`ZeitPunkt`) FROM `usdm` WHERE (`usdm`.`Pat_ID` = a.`Pat_id`))))));"
Call DtbCreateQueryDef(VN, Vsql)
vz = vz + 1


myEFrag ("DROP FUNCTION IF EXISTS `PatAlter`")
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `patalter`(pid INT(6) UNSIGNED) RETURNS SMALLINT(3) UNSIGNED" & vbCrLf & _
    "READS SQL DATA" & vbCrLf & _
    "COMMENT 'Patientenalter in Jahren'" & vbCrLf & _
"BEGIN" & vbCrLf & _
"DECLARE PAlter SMALLINT(3);" & vbCrLf & _
"SELECT TRUNCATE(DATEDIFF(NOW(),gebdat)/365.24,0) INTO PAlter FROM namen WHERE pat_id = pid;" & vbCrLf & _
"RETURN PAlter;" & vbCrLf & _
"END"
myEFrag (sql)

myEFrag ("DROP FUNCTION IF EXISTS `PatAltBr`")
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `PatAltBr`(pid INT(6) UNSIGNED) RETURNS FLOAT(5,2) UNSIGNED" & vbCrLf & _
    "READS SQL DATA" & vbCrLf & _
    "COMMENT 'Patientenalter in Jahren mit Bruch'" & vbCrLf & _
"BEGIN" & vbCrLf & _
"DECLARE PAlter FLOAT(5,2);" & vbCrLf & _
"SELECT TRUNCATE(DATEDIFF(NOW(),gebdat)/365.24,2) INTO PAlter FROM namen WHERE pat_id = pid;" & vbCrLf & _
"RETURN PAlter;" & vbCrLf & _
"END"
myEFrag (sql)

myEFrag ("DROP FUNCTION IF EXISTS `Metfdosis`")
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `Metfdosis`(pid INT(6) UNSIGNED) RETURNS int UNSIGNED" & vbCrLf & _
    "READS SQL DATA" & vbCrLf & _
    "COMMENT 'Metformindosis aus dem letzten Medikamentenplan'" & vbCrLf & _
"BEGIN" & vbCrLf & _
"DECLARE Metfd int;" & vbCrLf & _
"SELECT SUM(ROUND((zubruch(mo)+zubruch(mi)+zubruch(nm)+zubruch(ab)+zubruch(zn))*zunr(mp.medikament))) INTO Metfd " & vbCrLf & _
"FROM wmedplan mp " & vbCrLf & _
"LEFT JOIN medarten ma ON ma.medikament=mp.medanfang " & vbCrLf & _
"WHERE mp.pat_id=pid AND " & vbCrLf & _
"zeitpunkt=(SELECT MAX(zeitpunkt) FROM wmedplan WHERE pat_id=pid) AND metf " & vbCrLf & _
";" & vbCrLf & _
"IF ISNULL(Metfd) THEN SET Metfd=0; END IF;" & vbCrLf & _
"RETURN Metfd;" & vbCrLf & _
"END"
myEFrag (sql)

myEFrag ("DROP FUNCTION IF EXISTS `naemin`")
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `naemin`( pid INT(6) UNSIGNED, vmin DATETIME) RETURNS DATETIME " & vbCrLf & _
 "READS SQL DATA" & vbCrLf & _
 "COMMENT 'naechste freie Minute fuer Leistungseintrag' " & vbCrLf & _
"BEGIN" & vbCrLf & _
"  DECLARE nmin DATETIME DEFAULT STR_TO_DATE(DATE_FORMAT(vmin,'%Y%m%d%H%i'),'%Y%m%d%H%i');" & vbCrLf & _
"  DECLARE erg INT;" & vbCrLf & _
"runde:  LOOP" & vbCrLf & _
"  SET erg=0;" & vbCrLf & _
"  SELECT 1 INTO erg FROM leistungen WHERE pat_id=pid AND zeitpunkt=nmin LIMIT 1;" & vbCrLf & _
"  IF erg<>1 THEN LEAVE runde; END IF;" & vbCrLf & _
"  SET nmin=nmin+INTERVAL 1 minute;" & vbCrLf & _
"end loop;" & vbCrLf & _
"RETURN nmin;" & vbCrLf & _
"END"
myEFrag (sql)

myEFrag ("DROP FUNCTION IF EXISTS `zuUml`")
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `zuuml`(original VARCHAR(1000)) RETURNS VARCHAR(1000) " & _
      " DETERMINISTIC" & _
      " NO SQL " & vbCrLf & _
      " BEGIN " & vbCrLf & _
      "  DECLARE erg VARCHAR(1000); " & vbCrLf & _
      "  SET erg = REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( REPLACE( " & vbCrLf & _
      "                     REPLACE(original, 'Ä', 'Ae'), 'Ö', 'Oe'), 'Ü', 'Ue'), " & vbCrLf & _
      "                     'ä', 'ae'), 'ö', 'oe'), 'ü','ue'), 'ß', 'ss'); " & vbCrLf & _
      "  RETURN erg; " & vbCrLf & _
      " END"
    myEFrag (sql)
' sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `zuuml`(original VARCHAR(1000)) RETURNS VARCHAR(1000) CHARSET utf8mb4 COLLATE utf8mb4_german2_ci " & _

myEFrag ("DROP FUNCTION IF EXISTS `zuDatum`")
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `zuDatum`(original VARCHAR(1000), ausgang DATETIME) RETURNS DATE " & _
    "    DETERMINISTIC " & vbCrLf & _
    "    NO SQL " & vbCrLf & _
    "BEGIN " & vbCrLf & _
    " DECLARE temp,rest VARCHAR(1000);" & vbCrLf & _
    " DECLARE ch VARCHAR(1);" & vbCrLf & _
    " DECLARE i,njahr,posklammer,posw INT;" & vbCrLf & _
    " DECLARE muster VARCHAR(1000) DEFAULT '[0123456789]';" & vbCrLf & _
    " DECLARE tag, mon VARCHAR(2);" & vbCrLf & _
    " DECLARE jahr VARCHAR(5);" & vbCrLf & _
    " DECLARE Faktor DOUBLE;" & vbCrLf & _
    " DECLARE aktDATE, altDATE DATE;" & vbCrLf & _
    " DECLARE keinenummer BOOL;" & vbCrLf & _
    " IF ISNULL(original) OR original='' THEN RETURN DATE(ausgang); END IF;" & vbCrLf & _
    " aussen:   LOOP " & vbCrLf & _
    "  SET tag='01';" & vbCrLf & _
    "  SET mon='';" & vbCrLf & _
    "  SET i = 1;" & vbCrLf
sql = sql & _
    "  SET temp = '';" & vbCrLf & _
    "  SET Faktor=0;" & vbCrLf & _
    "  SET keinenummer=true;" & vbCrLf & _
    "loop_label:    LOOP " & Chr$(13) & _
    "   IF i > CHAR_LENGTH(original) THEN " & Chr$(13) & _
    "    LEAVE loop_label;" & Chr$(13) & _
    "   END IF;" & Chr$(13) & _
    "   SET ch = SUBSTRING(original,i,1);" & Chr$(13) & _
    "   IF ch REGEXP muster THEN " & Chr$(13) & _
    "    SET keinenummer=false;" & Chr$(13) & _
    "    IF LENGTH(temp)>3 THEN LEAVE loop_label; END IF;" & Chr$(13) & _
    "    SET temp = CONCAT(temp,ch);" & Chr$(13) & _
    "   ELSEIF ch = ',' THEN " & Chr$(13) & _
    "    IF INSTR(temp,'.')=0 THEN SET temp=CONCAT(temp,'.'); END IF;" & Chr$(13) & _
    "   ELSEIF ch REGEXP '[-/.]' THEN " & Chr$(13) & _
    "    IF INSTR(temp,'.') THEN LEAVE loop_label; END IF;" & Chr$(13) & _
    "    IF mon<>'' THEN SET tag = mon; END IF;" & Chr$(13) & _
    "    SET mon=LEFT(temp,2);" & Chr$(13) & _
    "    SET temp = '';" & Chr$(13) & _
    "   ELSEIF LENGTH(temp) > 0 THEN " & Chr$(13) & _
    "    LEAVE loop_label;" & Chr$(13) & _
    "   ELSE " & Chr$(13) & _
    "    SET temp = CONCAT(temp,'');" & Chr$(13) & _
    "   END IF;" & Chr$(13)
sql = sql & "   SET i=i+1;" & Chr$(13) & _
    "  END LOOP;" & Chr$(13) & _
    "  SET rest=ltrim(MID(original,i));" & Chr$(13) & _
    "  SET jahr = temp;" & Chr$(13) & _
    "  IF RIGHT(jahr,1)='.' THEN SET jahr = LEFT(jahr,LENGTH(jahr)-1); END IF;" & Chr$(13) & _
    "  IF Jahr='' THEN SET jahr='0'; END IF;" & Chr$(13) & _
    "  SET posklammer = INSTR(original,'(');" & Chr$(13) & _
    "  IF mon='' THEN " & Chr$(13) & _
    "    SET posw = INSTR(original,'Feb');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '2';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Winter');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '2';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'März');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '3';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'April');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '4';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Frühling');"
    sql = sql & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '4';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Mai');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '5';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Juni');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '6';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Juli');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '7';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Sommer');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '7';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'August');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '8';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Sept');"
    sql = sql & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '9';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Okt');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '10';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Herbst');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '10';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Nov');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '11';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Dez');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '12';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    IF mon = '' THEN SET mon = '1'; END IF;" & Chr$(13) & _
    "  END IF;" & Chr$(13) & _
    "  IF mon<>'1' THEN IF CAST(mon AS INT)>12 THEN SET mon='1'; END IF; END IF;" & Chr$(13) & _
    "  IF Jahr = 0 THEN SET Jahr = YEAR(ausgang); END IF;" & Chr$(13)
    sql = sql & _
    "  IF Jahr < 100 THEN " & Chr$(13) & _
    "    IF     rest LIKE 'wo%' THEN SET Faktor = 7;" & Chr$(13) & _
    "    ELSEIF rest LIKE 'mon%'  THEN SET Faktor = 30.437;" & Chr$(13) & _
    "    ELSEIF rest LIKE 'tag%'  THEN SET Faktor = 1;" & Chr$(13) & _
    "    ELSEIF rest LIKE 'a%' OR rest LIKE 'jah%' OR INSTR(jahr,'.') THEN SET Faktor = 365.24;" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    IF Faktor = 0 THEN " & Chr$(13) & _
    "     SET njahr = jahr + 2000;" & Chr$(13) & _
    "     IF njahr > YEAR(ausgang) THEN SET njahr=njahr-100; END IF;" & Chr$(13) & _
    "     IF mon='1' AND tag = '01' THEN " & Chr$(13) & _
    "      IF YEAR(ausgang) - Jahr > njahr THEN " & Chr$(13) & _
    "       SET njahr = YEAR(ausgang) - jahr;" & Chr$(13) & _
    "      END IF;" & Chr$(13) & _
    "     END IF;" & Chr$(13) & _
    "     SET aktDATE = DATE(CONCAT(njahr,'-',mon,'-',tag));" & Chr$(13) & _
    "    ELSE " & Chr$(13) & _
    "     SET aktDATE = DATE(DATE_SUB(ausgang, INTERVAL jahr * Faktor DAY));" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "  ELSE " & Chr$(13) & _
    "   SET aktDATE= DATE(CONCAT(jahr,'-',mon,'-',tag));" & Chr$(13) & _
    "  END IF;" & Chr$(13) & _
    "  IF keinenummer THEN " & Chr$(13) & _
    "   SET posw = INSTR(original,'heute');" & Chr$(13)
    sql = sql & _
    "   IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "    SET aktDATE = DATE(ausgang);" & Chr$(13) & _
    "    SET keinenummer = false;" & Chr$(13) & _
    "   END IF;" & Chr$(13) & _
    "   SET posw = INSTR(original,'gestern');" & Chr$(13) & _
    "   IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "    SET aktDATE = DATE(DATE_SUB(ausgang, INTERVAL 1 DAY));" & Chr$(13) & _
    "    SET keinenummer = false;" & Chr$(13) & _
    "   END IF;" & Chr$(13) & _
    "   SET posw = INSTR(original,'Jahre');" & Chr$(13) & _
    "   IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "    SET aktDATE = DATE(DATE_SUB(ausgang, INTERVAL 5 YEAR));" & Chr$(13) & _
    "    SET keinenummer = false;" & Chr$(13) & _
    "   END IF;" & Chr$(13) & _
    "   SET posw = INSTR(original,'Kindheit');" & Chr$(13) & _
    "   IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "    SET aktDATE = DATE(DATE_SUB(ausgang, INTERVAL 30 YEAR));" & Chr$(13) & _
    "    SET keinenummer = false;" & Chr$(13) & _
    "   END IF;" & Chr$(13) & _
    "  END IF;" & Chr$(13)
    sql = sql & _
    "  IF INSTR(rest,'(') THEN SET original=MID(rest,INSTR(rest,'(')+1); SET altDATE = aktDATE; ELSE LEAVE aussen; END IF;" & Chr$(13) & _
    "  END LOOP;" & Chr$(13) & _
    "  IF ISNULL(altDATE) THEN RETURN aktDATE; END IF;" & Chr$(13) & _
    "  IF not keinenummer AND aktDATE > altDATE THEN RETURN aktDATE; END IF;" & Chr$(13) & _
    "  RETURN altDATE;" & Chr$(13) & _
    " END"
    myEFrag (sql)
    
myEFrag ("DROP FUNCTION IF EXISTS `obrelDatum`")
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `obrelDatum`(original VARCHAR(1000), ausgang DATETIME) RETURNS Bool " & _
    "    DETERMINISTIC " & vbCrLf & _
    "    NO SQL " & vbCrLf & _
    "BEGIN " & vbCrLf & _
    " DECLARE temp VARCHAR(1000);" & vbCrLf & _
    " DECLARE ch VARCHAR(1);" & vbCrLf & _
    " DECLARE i INT;" & vbCrLf & _
    " DECLARE muster VARCHAR(1000) DEFAULT '[0123456789]';" & vbCrLf & _
    " DECLARE tag VARCHAR(2);" & vbCrLf & _
    " DECLARE mon VARCHAR(2);" & vbCrLf & _
    " DECLARE jahr VARCHAR(5);" & vbCrLf & _
    " DECLARE rest VARCHAR(1000);" & vbCrLf & _
    " DECLARE Faktor DOUBLE;" & vbCrLf & _
    " DECLARE njahr INTEGER;" & vbCrLf & _
    " DECLARE aktDATE DATE;" & vbCrLf & _
    " DECLARE altDATE DATE;" & vbCrLf & _
    "  DECLARE keinenummer bool;" & vbCrLf & _
    "  DECLARE posklammer INTEGER;" & vbCrLf & _
    "  DECLARE posw INTEGER;" & vbCrLf & _
    "  IF ISNULL(original) OR original='' THEN RETURN 0; END IF;" & vbCrLf & _
    " aussen:   LOOP " & vbCrLf & _
    "  SET tag='01';" & vbCrLf & _
    "  SET mon='';" & vbCrLf & _
    "  SET i = 1;" & vbCrLf
sql = sql & "  SET temp = '';" & vbCrLf & _
    "  SET Faktor=0;" & vbCrLf & _
    "  SET keinenummer=true;" & vbCrLf & _
    "loop_label:    LOOP " & Chr$(13) & _
    "   IF i > CHAR_LENGTH(original) THEN " & Chr$(13) & _
    "    LEAVE loop_label;" & Chr$(13) & _
    "   END IF;" & Chr$(13) & _
    "   SET ch = SUBSTRING(original,i,1);" & Chr$(13) & _
    "   IF ch REGEXP muster THEN " & Chr$(13) & _
    "    SET keinenummer=false;" & Chr$(13) & _
    "    IF LENGTH(temp)>3 THEN LEAVE loop_label; END IF;" & Chr$(13) & _
    "    SET temp = CONCAT(temp,ch);" & Chr$(13) & _
    "   ELSEIF ch = ',' THEN " & Chr$(13) & _
    "    IF INSTR(temp,'.')=0 THEN SET temp=CONCAT(temp,'.'); END IF;" & Chr$(13) & _
    "   ELSEIF ch REGEXP '[-/.]' THEN " & Chr$(13) & _
    "    IF INSTR(temp,'.') THEN LEAVE loop_label; END IF;" & Chr$(13) & _
    "    IF mon<>'' THEN SET tag = mon; END IF;" & Chr$(13) & _
    "    SET mon=LEFT(temp,2);" & Chr$(13) & _
    "    SET temp = '';" & Chr$(13) & _
    "   ELSEIF LENGTH(temp) > 0 THEN " & Chr$(13) & _
    "    LEAVE loop_label;" & Chr$(13) & _
    "   ELSE " & Chr$(13) & _
    "    SET temp = CONCAT(temp,'');" & Chr$(13) & _
    "   END IF;"
sql = sql & "   SET i=i+1;" & Chr$(13) & _
    "  END LOOP;" & Chr$(13) & _
    "  SET rest=ltrim(MID(original,i));" & Chr$(13) & _
    "  SET jahr = temp;" & Chr$(13) & _
    "  IF RIGHT(jahr,1)='.' THEN SET jahr = LEFT(jahr,LENGTH(jahr)-1); END IF;" & Chr$(13) & _
    "  IF Jahr='' THEN SET jahr='0'; END IF;" & Chr$(13) & _
    "  SET posklammer = INSTR(original,'(');" & Chr$(13) & _
    "  IF mon='' THEN " & Chr$(13) & _
    "    SET posw = INSTR(original,'Feb');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '2';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Winter');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '2';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'März');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '3';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'April');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '4';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Frühling');"
    sql = sql & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '4';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Mai');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '5';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Juni');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '6';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Juli');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '7';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Sommer');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '7';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'August');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '8';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Sept');"
    sql = sql & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '9';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Okt');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '10';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Herbst');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '10';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Nov');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '11';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    SET posw = INSTR(original,'Dez');" & Chr$(13) & _
    "    IF posw AND (posklammer = 0 OR posw < posklammer) THEN " & Chr$(13) & _
    "      SET mon = '12';" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    IF mon = '' THEN SET mon = '1'; END IF;" & Chr$(13) & _
    "  END IF;" & Chr$(13) & _
    "  IF mon<>'1' THEN IF CAST(mon AS INT)>12 THEN SET mon='1'; END IF; END IF;" & Chr$(13) & _
    "  IF Jahr < 100 THEN " & Chr$(13) & _
    "    IF     rest LIKE 'wo%' THEN SET Faktor = 7;"
    sql = sql & Chr$(13) & _
    "    ELSEIF rest LIKE 'mon%'  THEN SET Faktor = 30.437;" & Chr$(13) & _
    "    ELSEIF rest LIKE 'tag%'  THEN SET Faktor = 1;" & Chr$(13) & _
    "    ELSEIF rest LIKE 'a%' OR rest LIKE 'jah%' OR INSTR(jahr,'.') THEN SET Faktor = 365.24;" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "    IF Faktor = 0 THEN " & Chr$(13) & _
    "     SET njahr = jahr + 2000;" & Chr$(13) & _
    "     IF njahr > YEAR(ausgang) THEN SET njahr=njahr-100; END IF;" & Chr$(13) & _
    "     IF mon='1' AND tag = '01' THEN " & Chr$(13) & _
    "      IF YEAR(ausgang) - Jahr > njahr THEN " & Chr$(13) & _
    "       SET njahr = YEAR(ausgang) - jahr;" & Chr$(13) & _
    "      END IF;" & Chr$(13) & _
    "     END IF;" & Chr$(13) & _
    "     SET aktDATE = DATE(CONCAT(njahr,'-',mon,'-',tag));" & Chr$(13) & _
    "    ELSE " & Chr$(13) & _
    "     SET aktDATE = DATE(DATE_SUB(ausgang, INTERVAL jahr * Faktor DAY));" & Chr$(13) & _
    "    END IF;" & Chr$(13) & _
    "  ELSE " & Chr$(13) & _
    "   SET aktDATE= DATE(CONCAT(jahr,'-',mon,'-',tag));" & Chr$(13) & _
    "  END IF;" & Chr$(13) & _
    " END LOOP;" & Chr$(13) & _
    " return keinenummer;" & vbCrLf & _
    "END"
    myEFrag (sql)
    
    sql = "DROP FUNCTION IF EXISTS `zuNr`"
    myEFrag (sql)
    sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `zuNr`(istr VARCHAR(200)) RETURNS INT" & Chr$(13) & _
    "    DETERMINISTIC " & vbCrLf & _
    "    NO SQL " & Chr$(13) & _
    "BEGIN" & Chr$(13) & _
    "DECLARE INTi INTEGER DEFAULT 1;" & Chr$(13) & _
    "DECLARE sChar VARCHAR(1);" & Chr$(13) & _
    "DECLARE anfa INTEGER DEFAULT 0;" & Chr$(13) & _
    "DECLARE aufh INTEGER DEFAULT 0;" & Chr$(13) & _
    "DECLARE muster VARCHAR(13) DEFAULT '[0123456789]';" & Chr$(13) & _
    "DECLARE erg VARCHAR(200) DEFAULT '0';" & Chr$(13) & _
    "DECLARE maxl INTEGER(11) DEFAULT LENGTH(istr);"
    sql = sql & vbCrLf & _
    "WHILE INTi <= maxl AND NOT aufh DO" & Chr$(13) & _
    "SET sChar= SUBSTRING(istr,inti,1);" & Chr$(13) & _
    "IF sChar REGEXP muster THEN" & Chr$(13) & _
    " IF not anfa AND NOT aufh THEN SET anfa=1; END IF;" & Chr$(13) & _
    " ELSE" & Chr$(13) & _
    "  IF anfa THEN" & Chr$(13) & _
    "   SET aufh=1;" & Chr$(13) & _
    "  END IF;" & Chr$(13) & _
    " END IF;" & Chr$(13) & _
    "IF anfa AND NOT aufh THEN" & Chr$(13) & _
    " SET erg = CONCAT(erg,sChar);" & Chr$(13) & _
    "END IF;" & Chr$(13) & _
    "SET INTi=inti+1;" & Chr$(13) & _
    "END WHILE;" & Chr$(13) & _
    "RETURN CAST(erg AS UNSIGNED);" & Chr$(13) & "END "
    myEFrag (sql)
    
    sql = "DROP FUNCTION IF EXISTS `impfart`"
    myEFrag (sql)
    sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `impfart`(inh VARCHAR(400)) RETURNS int(3)" & Chr$(13) & _
    "    NO SQL DETERMINISTIC " & Chr$(13) & _
    "BEGIN" & Chr$(13) & _
    "DECLARE erg INTEGER DEFAULT 99;" & Chr$(13) & _
    "  CASE" & Chr$(13) & _
    "   WHEN inh RLIKE 'influvac|influsplit|vaxigrip|afluria|optaflu|xanaflu|begripal|begrivac|flucelvax|fluad|fluva|Flucelvax' THEN SET erg='1';" & vbCrLf & _
    "   WHEN inh RLIKE 'pneumovax|prevenar|apexxnar' THEN SET erg='2';" & vbCrLf & _
    "   WHEN inh RLIKE 'boostrix|repevax|revaxis|ipv[- ]M[eé]rieux|infanrix' THEN SET erg='3';" & vbCrLf & _
    "   WHEN inh RLIKE 'shingrix|zostavax' THEN SET erg='4';" & vbCrLf & _
    "   WHEN inh RLIKE 'fsme|encepur' THEN SET erg='5';" & vbCrLf & _
    "   WHEN inh RLIKE 'priorix|m-m-r|mmr' THEN SET erg='6';" & vbCrLf & _
    "   WHEN inh RLIKE 'td[- ]pur|tetagam|tetanol|td[- ]rix' THEN SET erg='7';" & vbCrLf & _
    "   WHEN inh RLIKE 'strovac' THEN SET erg='8';" & vbCrLf & _
    "   WHEN inh RLIKE 'havrix|twinrix|engerix b|VAQTA' THEN SET erg='9';" & vbCrLf & _
    "   WHEN inh RLIKE 'bexsero' THEN SET erg='10';" & vbCrLf & _
    "   WHEN inh RLIKE 'gardasil' THEN SET erg='12';" & vbCrLf & _
    "   WHEN inh RLIKE 'typhim' THEN SET erg='13';" & vbCrLf & _
    "   WHEN inh RLIKE 'Diphterie-Adsorbat' THEN SET erg='14';" & vbCrLf & _
    "   WHEN inh RLIKE 'rabipur' THEN SET erg='16';"
    sql = sql & _
    "   WHEN inh RLIKE 'vaixrip|grippeschutzimpfung|grippeimpfung|vaxogrip|grippeimpf.|inful|influvax|influfa|xanflu|optiflu|influvat|influrac|ifluvac|influspit|inflanrix|vagigrip|vaxigip|vaxgrip|vaxgrio|vaxigrio|vxigrp|infuvac|caxigrip|infusplit|inflvac|infliva|grippovax|flucelcax|flucelvac|grippovax|xanalfu|vaxigri |vaigrip|^axigrip|vaxugrip|influva |influcac|influvak|influcav|vacigrip|vaxigrim|invluvac|alfuria|influvasc|vaxipgrip|influfvac|indluvac|influavax|vxigrip|flucealvax|begropal|insuflac|saison tetra|efluelda' THEN SET erg='101';" & vbCrLf & _
    "   WHEN inh RLIKE 'pneumokokken|pneumokkoken|penumovax|pneumokken|pneumovac|pneumovx|pneumox|penumovac|pneumavac' THEN SET erg='102';" & vbCrLf & _
    "   WHEN inh RLIKE 'bosstrix|ipv ch|ipv mieurix|ipv meriueux|td- pur|bostrix' THEN SET erg='103';" & vbCrLf & _
    "   WHEN inh RLIKE 'fmse immun' THEN SET erg='105';" & vbCrLf & _
    "   WHEN inh RLIKE 'masern-impf.' THEN SET erg='106';" & vbCrLf & _
    "   WHEN inh RLIKE 'engerix-b' THEN SET erg='109';" & vbCrLf & _
    "   WHEN inh RLIKE 'bexero' THEN SET erg='110';" & vbCrLf & _
    "   WHEN inh RLIKE 'Janssen|Johnson' THEN SET erg='124';" & vbCrLf & _
    "   WHEN inh RLIKE 'Astra|Vaxzev' THEN SET erg='123';" & vbCrLf & _
    "   WHEN inh RLIKE 'Spikevax|Moderna' THEN SET erg='122';" & vbCrLf & _
    "   WHEN inh RLIKE 'biontech|comirnaty|corminaty|cominarty|comiarty|coirnaty|cominary|comirnary|commirnaty|comirnarty|cominaty|comitnaty' THEN IF inh RLIKE ': [GH]|: 23' THEN SET erg='127'; ELSE SET erg='121'; END IF;" & vbCrLf & _
    "   WHEN inh RLIKE 'Novavax' THEN SET erg='125';" & vbCrLf & _
    "   WHEN inh RLIKE 'Valneva' THEN SET erg='126';" & vbCrLf & _
    "   WHEN inh RLIKE 'VNR1M08G|Vit B12' THEN SET erg='199';" & vbCrLf & _
    "   ELSE SET erg='999';" & vbCrLf & _
    "  END CASE;" & vbCrLf & _
    " RETURN erg;" & vbCrLf & _
    "END"
    myEFrag (sql)
' die Chargen bei Covid fangen an mit G oder H
'    "   WHEN inh RLIKE 'covid' THEN SET erg='8';" & vbCrLf &

    sql = "DROP FUNCTION IF EXISTS `zuBruch`"
    myEFrag (sql)
    sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `zuBruch`(istr VARCHAR(200)) RETURNS DEC(11,2)" & Chr$(13) & _
    "    NO SQL DETERMINISTIC " & Chr$(13) & _
    "BEGIN" & Chr$(13) & _
    "DECLARE INTi INTEGER DEFAULT 1;" & Chr$(13) & _
    "DECLARE ric INT(1) DEFAULT 1;" & vbCrLf & _
    "DECLARE sChar VARCHAR(1);" & Chr$(13) & _
    "DECLARE anfa,aufh,teil,pz int(1) DEFAULT 0;" & Chr$(13) & _
    "DECLARE muster VARCHAR(18) DEFAULT '[-0123456789½¼/.,]';" & Chr$(13) & _
    "DECLARE erg,divi VARCHAR(200) DEFAULT '0';" & Chr$(13) & _
    "DECLARE ergnr,divnr DEC(11,2);" & Chr$(13) & _
    "DECLARE maxl INTEGER(11) DEFAULT LENGTH(istr);"
    sql = sql & vbCrLf & _
    "parse: WHILE INTi <= maxl AND NOT aufh DO" & Chr$(13) & _
    " SET sChar= SUBSTRING(istr,inti,1);" & Chr$(13) & _
    " IF sChar REGEXP muster THEN" & Chr$(13) & _
    "  IF NOT anfa AND NOT aufh THEN SET anfa=1; END IF;" & Chr$(13) & _
    " ELSEIF anfa THEN" & Chr$(13) & _
    "  SET aufh=1;" & Chr$(13) & _
    " END IF;" & Chr$(13) & _
    " IF anfa AND NOT aufh THEN" & Chr$(13) & _
    "  IF schar=',' THEN SET schar='.'; END IF;" & Chr$(13) & _
    "  IF schar='.' OR schar = '½' OR schar = '¼' THEN IF pz THEN LEAVE parse; ELSE SET pz=1; END IF; END IF;" & Chr$(13) & _
    "  CASE" & Chr$(13) & _
    "   WHEN schar='-' THEN IF ((teil AND divi='0') OR (teil=0 AND erg='0')) THEN SET ric=-ric; END IF;" & vbCrLf & _
    "   WHEN schar='/' THEN IF teil=1 THEN LEAVE parse; ELSE SET teil=1; END IF;" & Chr$(13) & _
    "   WHEN sChar='½' THEN SET erg='0.5';" & vbCrLf & _
    "   WHEN sChar = '¼' THEN SET erg='0.25';" & vbCrLf & _
    "   WHEN teil=0 THEN SET erg = CONCAT(erg,sChar);" & vbCrLf & _
    "   ELSE SET divi=CONCAT(divi,schar);" & vbCrLf & _
    "  END CASE;" & vbCrLf & _
    " END IF;" & vbCrLf & _
    " SET INTi=inti+1;" & vbCrLf & _
    "END WHILE;"
    sql = sql & vbCrLf & _
    "SET ergnr=CAST(erg AS dec(11,2));" & vbCrLf & _
    "IF teil THEN" & vbCrLf & _
    " SET divnr=CAST(divi AS dec(11,2));" & vbCrLf & _
    " IF divnr=0 THEN SET divnr=1; END IF;" & vbCrLf & _
    " RETURN ric*ergnr/divnr;" & vbCrLf & _
    "ELSE" & vbCrLf & _
    " RETURN ric*ergnr;" & vbCrLf & _
    "END IF;" & vbCrLf & _
    "END"
    myEFrag (sql)
    
' zu koordinieren mit doQuelldatum
myEFrag ("DROP FUNCTION IF EXISTS `quelldat`")
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `quelldat`(name VARCHAR(1000),dokd DATETIME) RETURNS DATETIME" & _
    "    DETERMINISTIC NO SQL" & Chr$(13) & _
    "BEGIN " & vbCrLf & _
    " DECLARE posf,runde,p1,art INT DEFAULT 0;" & vbCrLf & _
    " DECLARE Flb VARCHAR(10) DEFAULT 'Fremdlabor';" & vbCrLf & _
    " DECLARE m1 VARCHAR(1000) DEFAULT '(31\\.((0|)[13578]|1[02])|30\\.((0|)[1,3-9]|1[0-2])|((0|)[1-9]|[1-2][0-9])\\.((0|)[0-9]|1[0-2]))\\.(19|20|)([0-9]{2}) ([0-9]{6}|[0-9]{4})';" & vbCrLf & _
    " DECLARE m2 VARCHAR(1000) DEFAULT '(19|20)[0-9]{2}[01][0-9][0-3][0-9]_[0-2][0-9][0-6][0-9]{3}';" & vbCrLf & _
    " DECLARE m3 VARCHAR(1000) DEFAULT '(31\\.((0|)[13578]|1[02])|30\\.((0|)[1,3-9]|1[0-2])|((0|)[1-9]|[1-2][0-9])\\.((0|)[0-9]|1[0-2]))\\.(19|20|)([0-9]{2})( ([0-9]|[0-2][0-9])[.:][0-9]{2}([.:][0-9]{2}|)|)';" & vbCrLf & _
    " DECLARE m4 VARCHAR(1000) DEFAULT '(31\\.((0|)[13578]|1[02])|30\\.((0|)[1,3-9]|1[0-2])|((0|)[1-9]|[1-2][0-9])\\.((0|)[0-9]|1[0-2]))\\.';" & vbCrLf & _
    " DECLARE flname,testname,umwname VARCHAR(1000);" & vbCrLf & _
    " DECLARE eDAT DATETIME;" & vbCrLf & _
    " DECLARE EXIT HANDLER FOR SQLSTATE '22007' RETURN 18991230;" & vbCrLf & _
    " DECLARE EXIT HANDLER FOR SQLSTATE 'HY000' RETURN 18991230;" & vbCrLf
 sql = sql & _
    " SET POSf=INSTR(name,Flb);" & vbCrLf & _
    " IF POSf>0 THEN" & vbCrLf & _
    "  SET flname=SUBSTRING_INDEX(name,Flb,-1);" & vbCrLf & _
    "  SET p1=REGEXP_INSTR(flname,'BZ|RR|Medikamentenplan');" & vbCrLf & _
    "  IF p1>0 THEN SET flname=LEFT(flname,p1);" & vbCrLf & _
    "  END IF;" & vbCrLf & _
    " ELSE " & vbCrLf & _
    "  SET flname = name; " & vbCrLf & _
    " END IF;" & vbCrLf & _
    " na: LOOP" & vbCrLf & _
    "  SET art=0;" & vbCrLf & _
    "  SET testname=flname;" & vbCrLf
 sql = sql & _
    "  r1: LOOP" & vbCrLf & _
    "   SET p1=REGEXP_INSTR(testname,m1);" & vbCrLf & _
    "   IF p1>0 THEN SET art=1; SET umwname=REGEXP_SUBSTR(MID(testname,p1),m1);" & vbCrLf & _
    "   ELSE" & vbCrLf & _
    "    SET p1=REGEXP_INSTR(testname,m2);" & vbCrLf & _
    "    IF p1>0 THEN SET art=2; SET umwname=REGEXP_SUBSTR(MID(testname,p1),m2);" & vbCrLf & _
    "    ELSE" & vbCrLf & _
    "     SET p1=REGEXP_INSTR(testname,m3);" & vbCrLf & _
    "     IF p1>0 THEN SET art=3; SET umwname=REGEXP_SUBSTR(MID(testname,p1),m3);" & vbCrLf & _
    "     ELSE" & vbCrLf & _
    "      SET p1=REGEXP_INSTR(testname,m4);" & vbCrLf & _
    "      IF p1>0 THEN SET art=4; SET umwname=REGEXP_SUBSTR(MID(testname,p1),m4);" & vbCrLf & _
    "      END IF;" & vbCrLf & _
    "     END IF;" & vbCrLf & _
    "    END IF;" & vbCrLf & _
    "   END IF;" & vbCrLf & _
    "   IF p1=0 THEN LEAVE r1; END IF;" & vbCrLf & _
    "   SET flname=MID(testname,p1);" & vbCrLf & _
    "   SET testname=MID(flname,LENGTH(umwname)+1);" & vbCrLf & _
    "   IF REGEXP_INSTR(testname, '[0-9]') = 0 THEN LEAVE r1; END IF;" & vbCrLf & _
    "  END LOOP r1;" & vbCrLf
 sql = sql & _
    "  CASE WHEN art=1 THEN SET edat=STR_TO_DATE(REPLACE(umwname,':','.'),'%d.%m.%Y %H%i%s');" & vbCrLf & _
    "       WHEN art=2 THEN SET edat=STR_TO_DATE(umwname,'%Y%m%d_%H%i%s');" & vbCrLf & _
    "       WHEN art=3 THEN SET edat=STR_TO_DATE(REPLACE(REPLACE(umwname,':','.'),'-','.'),'%d.%m.%Y %H.%i.%s');" & vbCrLf & _
    "       WHEN art=4 THEN SET edat=STR_TO_DATE(CONCAT(REPLACE(umwname,':','.'),YEAR(dokd)),'%d.%m.%Y');" & vbCrLf & _
    "                       IF edat>dokd THEN SET edat=edat-INTERVAL 1 YEAR; END IF;" & vbCrLf & _
    "       ELSE SET edat=0;" & vbCrLf & _
    "  END CASE;" & vbCrLf & _
    "  IF runde=1 OR POSf=0 OR edat<>0 THEN LEAVE na; END IF;" & vbCrLf & _
    "  SET flname=name;" & vbCrLf & _
    "  SET runde=runde+1;" & vbCrLf & _
    " END LOOP na;" & vbCrLf & _
    " IF edat=0 OR edat>NOW() OR edat<19850101 THEN SET edat=18991230; END IF; -- (edat>dokd+INTERVAL 7 MONTH AND dokd<>18991230)" & vbCrLf & _
    " RETURN edat;" & vbCrLf & _
    "END" & vbCrLf
    myEFrag (sql)
'     "  " & vbCrLf & _


#If False Then
myEFrag ("DROP FUNCTION IF EXISTS `quelldat`")
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `quelldat`(name VARCHAR(1000)) RETURNS DATE " & _
    "    DETERMINISTIC " & _
    "    NO SQL " & Chr$(13) & _
    "BEGIN " & vbCrLf & _
    " DECLARE posf INT;" & vbCrLf & _
    " DECLARE Flb VARCHAR(10) DEFAULT 'Fremdlabor';" & vbCrLf & _
    " DECLARE muster VARCHAR(1000) DEFAULT '[0123456789]';" & vbCrLf & _
    " DECLARE erg VARCHAR(1000) DEFAULT '';" & vbCrLf & _
    " DECLARE ch VARCHAR(1);" & vbCrLf & _
    " DECLARE drw, arw, i INT DEFAULT 0; # Datum rückwärts (15.11.18/2018-11-15), Analyse (Namen) rückwärts" & vbCrLf & _
    " DECLARE tdat DATE DEFAULT 0;" & vbCrLf & _
    " DECLARE datum DATE DEFAULT 0;" & vbCrLf & _
    " DECLARE zifzahl INT DEFAULT 0;" & vbCrLf & _
    " DECLARE grupzahl INT DEFAULT 0;" & vbCrLf & _
    " DECLARE erwarte INT DEFAULT 0; # 1: Ziffer, 2: '.', 3 nichts mehr, 0: noch nichts" & vbCrLf & _
    " DECLARE endpos INT DEFAULT CHAR_LENGTH(name)+2;" & vbCrLf & _
    "# wenn 'Fremdlabor' vorkommt, das nä Datum danach nehmen" & vbCrLf & _
    " SET posf=INSTR(name,Flb);" & vbCrLf & _
    " IF POSf>0 THEN" & vbCrLf & _
    "  SET i=posf+CHAR_LENGTH(Flb);" & vbCrLf & _
    " END IF;" & vbCrLf & _
    " nochmal: LOOP" & vbCrLf & _
    "  infl: LOOP" & vbCrLf & _
    "    IF arw THEN SET i=i-1; ELSE SET i=i+1; END IF;" & vbCrLf & _
    "    IF i=endpos THEN LEAVE infl; END IF;" & vbCrLf
sql = sql & _
    "    SET ch = SUBSTRING(name,i,1);" & vbCrLf & _
    "    IF ch='.' AND erwarte>0 AND grupzahl<2 THEN" & vbCrLf & _
    "     SET grupzahl=grupzahl+1;" & vbCrLf & _
    "     SET zifzahl=0;" & vbCrLf & _
    "     SET erwarte=1;" & vbCrLf & _
    "     IF arw THEN SET erg=CONCAT(ch,erg); ELSE SET erg=CONCAT(erg,ch); END IF;" & vbCrLf & _
    "    ELSEIF ch REGEXP muster AND (erwarte=0 OR erwarte=1) THEN" & vbCrLf & _
    "     IF arw THEN SET erg=CONCAT(ch,erg); ELSE SET erg=CONCAT(erg,ch); END IF;" & vbCrLf & _
    "     SET zifzahl=zifzahl+1;" & vbCrLf & _
    "     IF grupzahl=2 AND ((arw=drw AND zifzahl=4) OR (arw<>drw AND zifzahl=2)) THEN" & vbCrLf & _
    "       SET erwarte=3;" & vbCrLf & _
    "     ELSEIF grupzahl<2 THEN" & vbCrLf & _
    "      IF ((arw<>drw AND zifzahl=4) OR (arw=drw AND zifzahl=2)) THEN SET erwarte=2; ELSE SET erwarte=1; END IF;" & vbCrLf & _
    "     END IF;" & vbCrLf & _
    "    ELSEIF erwarte>0 THEN" & vbCrLf & _
    "     SET tdat=STR_TO_DATE(erg,IF(drw,'%Y.%m.%d','%d.%m.%Y'));" & vbCrLf & _
    "     IF YEAR(tdat)>0 AND month(tdat)>0 AND day(tdat)>0 THEN" & vbCrLf & _
    "       SET datum=tdat;" & vbCrLf & _
    "       IF posf>0 THEN LEAVE infl; END IF; # ohne 'Fremdlabor' das letzte Datum nehmen" & vbCrLf & _
    "     END IF;" & vbCrLf & _
    "     SET erwarte=0;" & vbCrLf & _
    "     SET grupzahl=0;" & vbCrLf & _
    "     SET zifzahl=0;" & vbCrLf
sql = sql & _
    "     SET erg='';" & vbCrLf & _
    "    END IF;" & vbCrLf & _
    "  END LOOP infl;" & vbCrLf & _
    "  IF datum<>0 THEN LEAVE nochmal; END IF;" & vbCrLf & _
    "  IF name REGEXP '20[0123456789][0123456789]-[01][0123456789]-[0123][0123456789]' THEN" & vbCrLf & _
    "   SET name=REPLACE(name,'-','.');" & vbCrLf & _
    "   SET drw=1;" & vbCrLf & _
    "   SET erwarte=0; # evtl nicht nötig" & vbCrLf & _
    "   SET grupzahl=0; # evtl nicht nötig" & vbCrLf & _
    "   SET zifzahl=0; # evtl nicht nötig" & vbCrLf & _
    "   SET i=0;" & vbCrLf & _
    "   SET endpos=CHAR_LENGTH(name)+2;" & vbCrLf & _
    "# wenn nach Fremdlabor nichts gefunden, dann nochmal vorher suchen" & vbCrLf & _
    "  ELSEIF i=CHAR_LENGTH(name)+2 AND posf THEN" & vbCrLf & _
    "   SET i=posf+1;" & vbCrLf & _
    "   SET endpos=0;" & vbCrLf & _
    "   SET arw=1;" & vbCrLf & _
    "  ELSE" & vbCrLf & _
    "   LEAVE nochmal;" & vbCrLf & _
    "  END IF;" & vbCrLf & _
    " END LOOP; # nochmal;" & vbCrLf & _
    " RETURN datum;" & vbCrLf & _
    "END" & vbCrLf
    myEFrag (sql)
#End If

    ' quartal(zeitpunkt)=DiesQ() ist deutlich langsamer, 4.4.20
    VN = "konz"
    'Vsql = "SELECT pat_id, COUNT(DISTINCT DATE(zeitpunkt)) kz, fid FROM eintraege WHERE zeitpunkt BETWEEN qanf() AND qend() AND art IN (" & artspezG & ") GROUP BY pat_id; "
    Vsql = _
    "SELECT pat_id,fid, COUNT(DISTINCT zp) koz,GROUP_CONCAT(DISTINCT DATE_FORMAT(zp,'%e.%c.') ORDER BY zp SEPARATOR '•') czp " & vbCrLf & _
    ", MIN(zp) erst, MAX(zp) letzt " & vbCrLf & _
    " FROM ( " & vbCrLf & _
    " SELECT pat_id, CAST(ZeitPunkt As Date) zp,fid " & vbCrLf & _
    " FROM eintraege WHERE zeitpunkt BETWEEN DATE(CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY),'-',(QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+1,'-01')) AND CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY)+ quarter(NOW()-INTERVAL " & frist & " DAY) div 4 ,'-',((QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+4) mod 12,'-01')-INTERVAL 1 SECOND " & vbCrLf & _
    " UNION " & vbCrLf & _
    " SELECT pat_id, CAST(zeitpunkt As Date) zp,fid " & vbCrLf & _
    " FROM leistungen WHERE zeitpunkt BETWEEN DATE(CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY),'-',(QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+1,'-01')) AND CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY)+ quarter(NOW()-INTERVAL " & frist & " DAY) div 4 ,'-',((QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+4) mod 12,'-01')-INTERVAL 1 SECOND " & vbCrLf & _
    " UNION " & vbCrLf & _
    " SELECT pat_id, CAST(zeitpunkt As Date) zp,fid " & vbCrLf & _
    " FROM rezepteintraege WHERE zeitpunkt BETWEEN DATE(CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY),'-',(QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+1,'-01')) AND CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY)+ quarter(NOW()-INTERVAL " & frist & " DAY) div 4 ,'-',((QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+4) mod 12,'-01')-INTERVAL 1 SECOND " & vbCrLf & _
    ") i " & vbCrLf & _
    "GROUP BY pat_id;"
    Call DtbCreateQueryDef(VN, Vsql)
    vz = vz + 1
    
    VN = "kontkte" ' aktive Fälle mit Kontakten, schnell (ohne kateg und goäkatnr)
    ' Art: E=Eintrag, F=Formular, L=Leistung, M=Medplan
    ' quartal(e.zeitpunkt)= DiesQ()
    ' ktag fehlerhaft
    Vsql = _
    " SELECT dt,zp,pat_id,fid,schgr,vknr,ik,lanrid,GROUP_CONCAT(Typ ORDER BY Typ SEPARATOR '') Typ,GROUP_CONCAT(CONVERT(Art USING UTF8MB4) COLLATE utf8mb4_german2_ci ORDER BY Art) Art FROM ( " & vbCrLf & _
    "  SELECT DATE(e.zeitpunkt) dt, e.zeitpunkt zp,f.pat_id,f.fid,schgr,vknr,ik,f.lanrid, 'E' Typ, Art " & vbCrLf & _
    "   FROM aktfvs f " & vbCrLf & _
    "   JOIN eintraege e ON e.pat_id = f.pat_id " & vbCrLf & _
    "   AND e.zeitpunkt " & DiesQ() & vbCrLf & _
    "   AND art IN (" & artspezG & ") AND art<>'tv'" & vbCrLf & _
    "   GROUP BY pat_id,DATE(e.zeitpunkt) " & vbCrLf & _
    "  UNION " & vbCrLf & _
    "  SELECT DATE(k.zeitpunkt) dt, k.zeitpunkt zp,f.pat_id,f.fid,schgr,vknr,ik,f.lanrid, 'F' Typ, CONVERT(form_abk USING UTF8MB4) COLLATE utf8mb4_german2_ci Art " & vbCrLf & _
    "   FROM aktfvs f " & vbCrLf & _
    "   JOIN forminhkopf k ON k.pat_id = f.pat_id " & vbCrLf & _
    "   LEFT JOIN formulare u ON k.form_id=u.formid " & vbCrLf & _
    "   WHERE k.zeitpunkt " & DiesQ() & vbCrLf & _
    "   AND u.form_abk NOT LIKE '%dmp%' " & vbCrLf & _
    "   GROUP BY pat_id,DATE(k.zeitpunkt) " & vbCrLf
    ' ktag fehlerhaft
    Vsql = Vsql & _
    "  UNION " & vbCrLf & _
    "  SELECT DATE(l.zeitpunkt) dt, l.zeitpunkt zp,f.pat_id,f.fid,schgr,vknr,ik,f.lanrid, 'L' Typ, CAST(leistung AS CHAR COLLATE utf8mb4_german2_ci) Art " & vbCrLf & _
    "   FROM aktfvs f " & vbCrLf & _
    "   JOIN leistungen l ON l.pat_id = f.pat_id " & vbCrLf & _
    "   WHERE l.zeitpunkt " & DiesQ() & vbCrLf & _
    "   AND NOT leistung RLIKE '^32|^00000' " & vbCrLf & _
    "   GROUP BY pat_id,DATE(l.zeitpunkt) " & vbCrLf & _
    "  UNION " & vbCrLf & _
    "  SELECT DATE(m.zeitpunkt) dt, m.zeitpunkt zp,f.pat_id,f.fid,schgr,vknr,ik,f.lanrid, 'M' Typ, CONVERT('MP' USING UTF8MB4) COLLATE utf8mb4_german2_ci Art " & vbCrLf & _
    "   FROM aktfvs f " & vbCrLf & _
    "   JOIN medplan m ON m.pat_id = f.pat_id " & vbCrLf & _
    "   WHERE m.zeitpunkt " & DiesQ() & vbCrLf & _
    "   GROUP BY pat_id,DATE(m.zeitpunkt) " & vbCrLf & _
    "  ) i GROUP BY pat_id,dt " & vbCrLf
    ' quartal(..) = DiesQ() 29s ggü 55s mit .. zeitpunkt BETWEEN qanf() AND qend()
    ' "   AND (leistung='00000' OR leistung LIKE '32%') " & vbCrLf & ' => genau gleich schnell
    Call DtbCreateQueryDef(VN, Vsql)
    vz = vz + 1

    VN = "kontktep" ' aktive Fälle mit Kontakten, schnell (ohne kateg und goäkatnr)
    ' Art: E=Eintrag, F=Formular, L=Leistung, M=Medplan
    ' quartal(e.zeitpunkt)= DiesQ()
    ' ktag fehlerhaft
    ' eintraege ausgewählt s. kstreng
    Vsql = _
    " SELECT dt,zp,pat_id,fid,schgr,vknr,ik,lanrid,GROUP_CONCAT(Typ ORDER BY Typ SEPARATOR '') Typ,GROUP_CONCAT(Art ORDER BY Art) Art FROM ( " & vbCrLf & _
    "  SELECT DATE(e.zeitpunkt) dt, e.zeitpunkt zp,f.pat_id,f.fid,schgr,vknr,ik,f.lanrid, 'E' Typ, Art " & vbCrLf & _
    "   FROM aktfvs f " & vbCrLf & _
    "   JOIN eintraege e ON e.pat_id = f.pat_id " & vbCrLf & _
    "   AND e.zeitpunkt " & DiesQ() & vbCrLf & _
    "   AND art IN (" & artSpezBerat & "," & artSpezEintr & "," & artSpezUS1 & ") AND NOT art IN ('re','ba','doppler','duplex','impf','inj','kva','tv','ufrag','wv','caro','colo','debr','fa','EKG','GPD','kv','LZRR','OAU','pa','pros','puls','rp','sono','ulcus','vac','ADL'," & artSpezMA & ") AND art<>'tv'" & vbCrLf & _
    "   GROUP BY pat_id,DATE(e.zeitpunkt) " & vbCrLf & _
    "  ) i GROUP BY pat_id,dt " & vbCrLf
    ' quartal(..) = DiesQ() 29s ggü 55s mit .. zeitpunkt BETWEEN qanf() AND qend()
    ' "   AND (leistung='00000' OR leistung LIKE '32%') " & vbCrLf & ' => genau gleich schnell
    Call DtbCreateQueryDef(VN, Vsql)
    vz = vz + 1
    
    VN = "aktfkvs" ' aktive Fälle mit Kontakten, auch 0 Kontakte, schnell (ohne kateg und goäkatnr)
    Vsql = _
    "SELECT f.pat_id, COUNT(DISTINCT zp) koz,quartal,COALESCE(GROUP_CONCAT(DISTINCT DATE_FORMAT(zp,'%e.%c.') ORDER BY zp SEPARATOR '•'),'') czp, COALESCE(min(zp),'') erst, COALESCE(MAX(zp),'') letzt, f.fid,f.schgr,f.vknr,f.ik,f.lanrid, COALESCE(GROUP_CONCAT(Typ ORDER BY zp SEPARATOR '•'),'') Typ, COALESCE(GROUP_CONCAT(Art ORDER BY zp SEPARATOR '•'),'') Art" & vbCrLf & _
    "FROM aktfv f " & vbCrLf & _
    "LEFT JOIN kontkte USING (pat_id)" & vbCrLf & _
    "GROUP BY f.pat_id; "
    Call DtbCreateQueryDef(VN, Vsql)
    vz = vz + 1

    VN = "aktfpkvs" ' aktive Fälle mit Kontakten, auch 0 Kontakte, schnell (ohne kateg und goäkatnr)
    Vsql = _
    "SELECT f.pat_id, COUNT(DISTINCT zp) koz,COALESCE(GROUP_CONCAT(DISTINCT DATE_FORMAT(zp,'%e.%c.') ORDER BY zp SEPARATOR '•'),'') czp, COALESCE(min(zp),'') erst, COALESCE(MAX(zp),'') letzt, f.fid,f.schgr,f.vknr,f.ik,f.lanrid, COALESCE(GROUP_CONCAT(Typ ORDER BY zp SEPARATOR '•'),'') Typ, COALESCE(GROUP_CONCAT(Art ORDER BY zp SEPARATOR '•'),'') Art" & vbCrLf & _
    "FROM aktfv f " & vbCrLf & _
    "LEFT JOIN kontktep USING (pat_id)" & vbCrLf & _
    "GROUP BY f.pat_id; "
    Call DtbCreateQueryDef(VN, Vsql)
    vz = vz + 1

    VN = "kstreng" ' aktive Fälle mit Kontakten, auch 0 Kontakte, schnell (ohne kateg und goäkatnr)
    Vsql = _
    " SELECT zp, GROUP_CONCAT(was SEPARATOR '; ') was, pat_id,fid,schgr,vknr,ik,lanrid  FROM ( " & vbCrLf & _
    "  SELECT m.zeitpunkt zp, 'Medplan erstellt' was, f.* " & vbCrLf & _
    "   FROM aktfvs f " & vbCrLf & _
    "   LEFT JOIN medplan m ON m.pat_id = f.pat_id " & vbCrLf & _
    "   WHERE m.zeitpunkt BETWEEN DATE(CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY),'-',(QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+1,'-01')) AND CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY)+ quarter(NOW()-INTERVAL " & frist & " DAY) div 4 ,'-',((QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+4) mod 12,'-01')-INTERVAL 1 SECOND " & vbCrLf & _
    "   GROUP BY pat_id,DATE(m.zeitpunkt) " & vbCrLf & _
    "  UNION ALL " & vbCrLf & _
    "  SELECT e.zeitpunkt zp, GROUP_CONCAT(CONCAT(art,' ',inhalt) SEPARATOR '; ') was, f.*" & vbCrLf & _
    "   FROM aktfvs f " & vbCrLf & _
    "   LEFT JOIN eintraege e ON e.pat_id = f.pat_id " & vbCrLf & _
    "   AND e.zeitpunkt BETWEEN DATE(CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY),'-',(QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+1,'-01')) AND CONCAT(YEAR(NOW()-INTERVAL " & frist & " DAY)+ quarter(NOW()-INTERVAL " & frist & " DAY) div 4 ,'-',((QUARTER(NOW()-INTERVAL " & frist & " DAY)-1)*3+4) mod 12,'-01')-INTERVAL 1 SECOND AND art IN (" & artSpezBerat & "," & artSpezEintr & "," & artSpezUS1 & ") AND NOT art IN ('re','ba','doppler','duplex','impf','inj','kva','tv','ufrag','wv','caro','colo','debr','fa','EKG','GPD','kv','LZRR','OAU','pa','pros','puls','rp','sono','ulcus','vac','ADL'," & artSpezMA & ") " & vbCrLf & _
    "   GROUP BY pat_id,DATE(e.zeitpunkt) " & vbCrLf & _
    "  ) i GROUP BY pat_id,DATE(zp) " & vbCrLf
    Call DtbCreateQueryDef(VN, Vsql)
    vz = vz + 1
    
    sql = "DROP FUNCTION IF EXISTS `qende`"
    myEFrag (sql)
    sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `qende`(datum DATETIME) RETURNS DATE" & _
    "    DETERMINISTIC" & vbCrLf & _
    "    NO SQL" & vbCrLf & _
    "    COMMENT 'Quartalsende zu bestimmtem Datum' " & _
    "BEGIN " & vbCrLf & _
    " DECLARE erg DATE; " & vbCrLf & _
    " SET erg = SUBDATE(ADDDATE(qbeg(DATE(CONCAT(YEAR(datum),'-',(QUARTER(datum)-1)*3+1,'-01'))), INTERVAL 3 MONTH),INTERVAL 1 DAY);  RETURN erg;" & vbCrLf & _
    "END "
    myEFrag (sql)

'    sql = "DROP FUNCTION IF EXISTS `artspezg`"
'    myEFrag (sql)
'    sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `artspezg`() RETURNS VARCHAR(10000)" & Chr$(13) & _
'    "    DETERMINISTIC" & Chr$(13) & _
'    "BEGIN " & Chr$(13) & _
'    " RETURN " & Chr(34) & "(" & artspezG & ")" & Chr(34) & ";" & Chr$(13) & _
'    "END "
'    myEFrag (sql)

myEFrag ("DROP FUNCTION IF EXISTS `quelle`.`xpneuzuypneu`")
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `quelle`.`xpneuzuypneu`(xid INT) RETURNS INT " & _
    "READS SQL DATA " & _
    "BEGIN " & vbCrLf & _
"DECLARE yid INT; " & vbCrLf & _
"SELECT y.id INTO yid " & vbCrLf & _
"FROM (SELECT x.*,labor FROM laborxpneu x LEFT JOIN laborxplab lx ON x.lid=lx.id) x " & vbCrLf & _
"LEFT JOIN (SELECT y.*,ly.labor FROM laborypneu y LEFT JOIN laboryplab ly ON y.labid=ly.id) y ON y.abkü=x.abkü AND y.langtext=x.langtext AND (y.einheit=x.einheit OR ( " & vbCrLf & _
"(SELECT MAX(yi.id) FROM (SELECT yii.*,lyii.labor FROM laborypneu yii LEFT JOIN laboryplab lyii ON yii.labid=lyii.id) yi WHERE yi.abkü=x.abkü AND yi.langtext=x.langtext AND yi.einheit=x.einheit AND yi.labor=x.labor) IS NULL AND " & vbCrLf & _
"x.einheit='kA') " & vbCrLf & _
") AND y.labor=x.labor " & vbCrLf & _
"WHERE x.id=xid; " & vbCrLf & _
"RETURN yid;" & vbCrLf & _
"END;" & vbCrLf
myEFrag (sql)

' 25.5.14: auch die Funktionen:
sql = "DROP FUNCTION IF EXISTS `quelle`.`gesname`"
myEFrag (sql)
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION  `quelle`.`gesname`( id INT(6) UNSIGNED ) RETURNS VARCHAR(26) CHARSET utf8mb4 " & _
    "READS SQL DATA " & _
    "DETERMINISTIC " & _
    "COMMENT 'gesname für Abrechnungsprogramme' " & _
"BEGIN " & vbCrLf & _
"DECLARE gesname VARCHAR(26) CHARACTER SET utf8; " & vbCrLf & _
"SELECT LEFT(CONCAT(IF(n.titel='','',CONCAT(n.titel,' ')),IF(n.nvorsatz='','',CONCAT(n.nvorsatz,' ')),n.nachname,', ',n.vorname),25) Name INTO gesname FROM namen n WHERE pat_id = id; " & vbCrLf & _
"RETURN gesname; " & vbCrLf & _
"END; "
myEFrag (sql)

sql = "DROP FUNCTION IF EXISTS `quelle`.`gesnameg`"
myEFrag (sql)
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION  `quelle`.`gesnameg`( id INT(6) UNSIGNED ) RETURNS VARCHAR(40) CHARSET utf8mb4 " & _
    "READS SQL DATA " & _
    "DETERMINISTIC " & _
    "COMMENT 'gesname für Abrechnungsprogramme' " & _
"BEGIN " & vbCrLf & _
"DECLARE gesname VARCHAR(40) CHARACTER SET utf8; " & vbCrLf & _
"SELECT CONCAT(LEFT(CONCAT(IF(n.titel='','',CONCAT(n.titel,' ')),IF(n.nvorsatz='','',CONCAT(n.nvorsatz,' ')),n.nachname,', ',n.vorname),25),', * ',DATE_FORMAT(n.gebdat,'%e.%c.%Y')) Name INTO gesname FROM namen n WHERE pat_id = id;" & vbCrLf & _
"RETURN gesname; " & vbCrLf & _
"END; "
myEFrag (sql)

sql = "DROP FUNCTION IF EXISTS `quelle`.`anteil`"
myEFrag sql
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `anteil`(orig text, teiler INTEGER) RETURNS DECIMAL(2,1) " & _
      " DETERMINISTIC" & vbCrLf & _
      " NO SQL" & vbCrLf & _
      " COMMENT 'Monofilament und Vibration' " & _
      " BEGIN " & vbCrLf & _
      " DECLARE links, linksneu, rechts, rechtsneu text; " & vbCrLf & _
      " DECLARE i, pos INT; " & vbCrLf & _
      " DECLARE ch VARCHAR(1); " & vbCrLf & _
      " DECLARE erg DECIMAL(2,1); " & vbCrLf & _
      " SET pos=INSTR(orig,'/'); " & vbCrLf & _
      " IF pos=0 THEN " & vbCrLf & _
      "  IF orig REGEXP '[-0-9]+' THEN " & vbCrLf & _
      "   SET linksneu=''; " & vbCrLf & _
      "   SET i = 1; " & vbCrLf & _
      "   loop_0:  LOOP " & vbCrLf & _
      "    IF i>CHAR_LENGTH(orig) THEN LEAVE loop_0; END IF; " & vbCrLf & _
      "    SET ch = SUBSTRING(orig,i,1); " & vbCrLf & _
      "    IF ch='-' AND linksneu='' THEN SET linksneu='0'; END IF; " & vbCrLf & _
      "    IF ch='>' THEN SET linksneu='0'; LEAVE loop_0; END IF; " & vbCrLf & _
      "    IF INSTR('-(?,',ch) AND linksneu <> ''  THEN LEAVE loop_0; END IF; " & vbCrLf & _
      "    IF ch REGEXP '[0-9]' THEN SET linksneu = CONCAT(linksneu,ch); END IF; " & vbCrLf & _
      "    IF linksneu<>'' THEN LEAVE loop_0; END IF; " & vbCrLf & _
      "    SET i=i+1; " & vbCrLf
 sql = sql & _
      "   END LOOP; " & vbCrLf & _
      "   IF teiler=0 THEN" & vbCrLf & _
      "    SET erg = -1;" & vbCrLf & _
      "   ELSE" & vbCrLf & _
      "    SET erg = linksneu/teiler; " & vbCrLf & _
      "   END IF;" & vbCrLf & _
      "  ELSE " & vbCrLf & _
      "   SET erg=-1; " & vbCrLf & _
      "  END IF; " & vbCrLf & _
      " ELSE " & vbCrLf
 sql = sql & _
      "  SET links=SUBSTRING(orig,1,pos-1); " & vbCrLf & _
      "  SET rechts=SUBSTRING(orig,pos+1); " & vbCrLf & _
      "  IF links REGEXP '[-0-9]+' AND NOT links REGEXP '^[0-9]+$' THEN " & vbCrLf & _
      "   SET linksneu=''; " & vbCrLf & _
      "   SET i = 1; " & vbCrLf & _
      "   loop_label:    LOOP " & vbCrLf & _
      "   IF i>CHAR_LENGTH(orig) THEN LEAVE loop_label; END IF; " & vbCrLf & _
      "   SET ch = SUBSTRING(links,i,1); " & vbCrLf & _
      "   IF ch='-' AND linksneu='' THEN SET linksneu='0'; END IF; " & vbCrLf & _
      "   IF ch='>' THEN SET linksneu='0'; LEAVE loop_label; END IF; " & vbCrLf & _
      "   IF INSTR('-(?,',ch) AND linksneu <> '' THEN LEAVE loop_label; END IF; " & vbCrLf & _
      "   IF ch REGEXP '[0-9]' THEN SET linksneu = CONCAT(linksneu,ch); END IF; " & vbCrLf & _
      "   IF linksneu<>'' THEN LEAVE loop_label; END IF; " & vbCrLf & _
      "   SET i=i+1; " & vbCrLf & _
      "  END LOOP; " & vbCrLf & _
      "  SET links = linksneu; " & vbCrLf & _
      " END IF; " & vbCrLf & _
      " IF rechts REGEXP '[0-9]+' AND NOT rechts REGEXP '^[0-9]+$' THEN " & vbCrLf & _
      "  SET rechtsneu=''; " & vbCrLf & _
      "  SET i = 1; " & vbCrLf & _
      "  loop_1:  LOOP " & vbCrLf & _
      "   IF i>CHAR_LENGTH(orig) THEN LEAVE loop_1; END IF; " & vbCrLf & _
      "   SET ch = SUBSTRING(rechts,i,1); " & vbCrLf & _
      "   IF INSTR('-(?,',ch) THEN LEAVE loop_1; END IF; " & vbCrLf
sql = sql & _
      "   IF ch REGEXP '[0-9]' THEN SET rechtsneu = CONCAT(rechtsneu,ch); END IF; " & vbCrLf & _
      "   IF rechtsneu<>'' THEN LEAVE loop_1; END IF; " & vbCrLf & _
      "   SET i=i+1; " & vbCrLf & _
      "  END LOOP; " & vbCrLf & _
      "  SET rechts = rechtsneu; " & vbCrLf & _
      " END IF; " & vbCrLf
sql = sql & _
      " IF links REGEXP '[0-9]+' AND rechts REGEXP '[0-9]+' THEN " & vbCrLf & _
      "  IF rechts=0 THEN" & vbCrLf & _
      "   SET erg = -1;" & vbCrLf & _
      "  ELSE" & vbCrLf & _
      "   SET erg = links/rechts; " & vbCrLf & _
      "  END IF;" & vbCrLf & _
      " ELSEIF links REGEXP '[0-9]+' THEN " & vbCrLf & _
      "  IF teiler=0 THEN" & vbCrLf & _
      "   SET erg = -1;" & vbCrLf & _
      "  ELSE" & vbCrLf & _
      "   SET erg = links/teiler; " & vbCrLf & _
      "  END IF;" & vbCrLf & _
      " ELSE " & vbCrLf & _
      "  SET erg=-1; " & vbCrLf & _
      " END IF; " & vbCrLf & _
      "END IF; " & vbCrLf & _
      "RETURN erg; " & vbCrLf & _
      "END; "
myEFrag (sql)

'Const problematisch% = True
#If problematisch Then
    sql = "DROP PROCEDURE IF EXISTS `quelle`.`geslabp`"
    myEFrag sql
    sql = _
    "CREATE DEFINER=`praxis`@`%` PROCEDURE `geslabp`(in pid INT(6)) " & vbCrLf & _
    "    READS SQL DATA " & vbCrLf & _
    "    COMMENT 'Labor zu einem Patienten' " & vbCrLf & _
    "BEGIN " & vbCrLf & _
    " SELECT * FROM (" & _
    "  SELECT * FROM labor2a WHERE pat_id = pid " & vbCrLf & _
    "  UNION " & vbCrLf & _
    "  SELECT * FROM labor1a WHERE pat_id = pid" & vbCrLf & _
    " ) i GROUP BY zeitpunkt,abkü,wert ORDER BY zeitpunkt DESC,abkü,wert,einheit,nb; " & vbCrLf & _
    "END;"
    myEFrag (sql)
    
    sql = "DROP PROCEDURE IF EXISTS `quelle`.`geslabap`"
    myEFrag sql
    sql = _
    "CREATE DEFINER=`praxis`@`%` PROCEDURE `geslabap`(in pid INT(6)) " & vbCrLf & _
    "    READS SQL DATA " & vbCrLf & _
    "    COMMENT 'Labor zu einem Patienten' " & vbCrLf & _
    "BEGIN " & vbCrLf & _
    "SELECT GROUP_CONCAT(wert SEPARATOR '<br>') wert,abkü, einheit,Zeitpunkt,NB,gruppe,reihe,uNg,oNg,Pfad,Kommentar,Labor,Langtext,pat_id " & vbCrLf & _
    " FROM (SELECT * FROM " & vbCrLf & _
    "  (SELECT wert,abkü,einheit,zeitpunkt,nb,gruppe,reihe,uNg,oNg,Pfad,Kommentar,Labor,Langtext,pat_id FROM labor2a l2 WHERE pat_id=pid " & vbCrLf & _
    "    UNION " & vbCrLf & _
    "   SELECT wert,abkü,einheit,zeitpunkt,nb,gruppe,reihe,uNg,oNg,Pfad,Kommentar,Labor,Langtext,pat_id FROM labor1a l1 WHERE pat_id=pid) i " & vbCrLf & _
    "  GROUP BY zeitpunkt,abkü,wert) i " & vbCrLf & _
    " GROUP BY abkü,einheit,nb,zeitpunkt" & vbCrLf & _
    "ORDER BY zeitpunkt DESC,abkü,wert,einheit,nb;" & vbCrLf & _
    "END;"
    myEFrag (sql)
#End If ' problematisch

sql = "DROP PROCEDURE IF EXISTS `quelle`.`geslabdp`"
myEFrag sql
sql = _
"CREATE DEFINER=`praxis`@`%` PROCEDURE `geslabdp`(IN pid INT(6),IN esql VARCHAR(1000)) " & vbCrLf & _
"    READS SQL DATA " & vbCrLf & _
"    COMMENT 'Labor zu einem Patienten' " & vbCrLf & _
"BEGIN " & vbCrLf & _
"DECLARE rest VARCHAR(100) DEFAULT ',nb,gruppe,reihe,uNg,oNg,Pfad,Kommentar,Labor,Langtext,pat_id,DatID ';" & vbCrLf & _
"SET @frag = CONCAT(""" & vbCrLf & _
"SELECT COALESCE(GROUP_CONCAT(DISTINCT CASE WHEN wert<>'' THEN wert END SEPARATOR '<br>'),'') Wert, abkü, einheit,Zeitpunkt,COUNT(DISTINCT abkü,einheit,zeitpunkt) Zl"",rest,"" FROM " & vbCrLf & _
" (SELECT wert,abkü,einheit,zeitpunkt"",rest,"" FROM labor2a WHERE pat_id="",pid,""" & vbCrLf & _
"  UNION " & vbCrLf & _
"  SELECT wert,abkü,einheit,zeitpunkt"",rest,"" FROM labor1a l1 WHERE pat_id="",pid,"") i " & vbCrLf & _
" "",IF(esql='','GROUP BY zeitpunkt DESC,abkü,einheit;',esql));" & vbCrLf & _
"PREPARE sc1 FROM @frag; " & vbCrLf & _
"EXECUTE sc1;" & vbCrLf & _
"DEALLOCATE PREPARE sc1; " & vbCrLf & _
"-- SELECT @frag; " & vbCrLf & _
"END;"
myEFrag (sql)

sql = "DROP PROCEDURE IF EXISTS `quelle`.`getfeldinhvw`"
myEFrag sql
sql = _
"CREATE DEFINER=`praxis`@`%` PROCEDURE `getfeldinhvw`(IN `felder` LONGTEXT, IN `stbyte` INTEGER) " & vbCrLf & _
"    LANGUAGE SQL" & vbCrLf & _
"    NOT DETERMINISTIC" & vbCrLf & _
"    MODIFIES SQL DATA" & vbCrLf & _
"    SQL SECURITY DEFINER" & vbCrLf & _
"    COMMENT ''" & vbCrLf & _
"BEGIN" & vbCrLf & _
" DECLARE pos1,pos2,aktvw,iru INTEGER DEFAULT 1;" & vbCrLf & _
" DECLARE elem,vw LONGTEXT;" & vbCrLf & _
" START TRANSACTION;" & vbCrLf & _
" SET vw='';" & vbCrLf & _
"begparse:  LOOP" & vbCrLf & _
"  SET pos2=LOCATE('^|',felder,pos1);" & vbCrLf & _
"  IF pos2=0 THEN SET elem=MID(felder,pos1); ELSE" & vbCrLf & _
"  SET elem=mid(felder,pos1,pos2-pos1); END IF;" & vbCrLf & _
"  SET aktvw=0;" & vbCrLf
sql = sql & _
"  SET iru=1;" & vbCrLf & _
"irulab: LOOP" & vbCrLf & _
"  SELECT COALESCE((SELECT feldinhvw FROM forminhaltfeldinh WHERE feldinh = CONVERT(elem,BINARY) LIMIT 1),0) INTO aktvw;" & vbCrLf & _
"   IF aktvw = 0 THEN" & vbCrLf & _
"    INSERT INTO forminhaltfeldinh(feldinh,stbyte) VALUES(CONVERT(elem,BINARY),stbyte);" & vbCrLf & _
"--   SET aktvw=last_insert_id();" & vbCrLf & _
"   END IF;" & vbCrLf & _
"   IF aktvw<>0 OR iru=2 THEN LEAVE irulab; END IF;" & vbCrLf & _
"   SET iru=iru+1;" & vbCrLf & _
"  END LOOP irulab;" & vbCrLf & _
"  IF vw<>'' THEN SET vw=concat(vw,','); END IF;" & vbCrLf & _
"  SET vw=concat(vw,aktvw);" & vbCrLf & _
"  IF pos2=0 THEN LEAVE begparse; END IF;" & vbCrLf & _
"  SET pos1=pos2+2;" & vbCrLf & _
" END LOOP begparse;" & vbCrLf & _
" COMMIT;" & vbCrLf & _
" SET @vw=vw;" & vbCrLf & _
" SELECT vw;" & vbCrLf & _
"END;"
myEFrag (sql)

' FROM laboryus u LEFT JOIN laborywert w ON w.usid=u.id LEFT JOIN laboryhinw e ON e.erklid=w.id LEFT JOIN laboryhinw k ON k.kommid=w.id LEFT JOIN laborypnb n ON n.nbid=w.id LEFT JOIN laborysaetze s ON s.satzid=u.satzid LEFT JOIN laborydat d ON d.datid=s.datid LEFT JOIN laboryplab l ON l.labid=s.id LEFT JOIN laborparameter p ON p.abkü=w.abkü AND p.einheit=IF(w.einheit IN ('','\'kA\''),'kA',w.einheit) HAVING (wert<>'' AND wert IS NOT NULL) OR (kommentar<>'' AND kommentar IS NOT NULL)


sql = "DROP PROCEDURE IF EXISTS `quelle`.`geslabkatz`"
myEFrag sql
sql = _
"CREATE DEFINER=`praxis`@`%` PROCEDURE `geslabkatz`(in pid INT(6),in katg VARCHAR(20)) " & vbCrLf & _
"    READS SQL DATA " & vbCrLf & _
"    COMMENT 'Zahl der verschiedenen Labortage oder -Abkürzungen oder usw.' " & vbCrLf & _
"BEGIN " & vbCrLf & _
"SET @frag = CONCAT('" & vbCrLf & _
"SELECT COUNT(0) Zahl FROM " & vbCrLf & _
" (SELECT ',katg,' FROM labor2a WHERE pat_id=',pid,'" & vbCrLf & _
"  UNION " & vbCrLf & _
"  SELECT ',katg,' FROM labor1a l1 WHERE pat_id=',pid,') i');" & vbCrLf & _
"PREPARE sc1 FROM @frag; " & vbCrLf & _
"execute sc1;" & vbCrLf & _
"DEALLOCATE PREPARE sc1; " & vbCrLf & _
"-- SELECT @frag; " & vbCrLf & _
"END;"
myEFrag (sql)

sql = "DROP PROCEDURE IF EXISTS `quelle`.`geslabkatg`"
myEFrag sql
sql = _
"CREATE DEFINER=`praxis`@`%` PROCEDURE `geslabkatg`(in pid INT(6),in katg VARCHAR(20)) " & vbCrLf & _
"    READS SQL DATA " & vbCrLf & _
"    COMMENT 'Zahl der verschiedenen Labortage' " & vbCrLf & _
"BEGIN " & vbCrLf & _
"SET @frag = CONCAT('" & vbCrLf & _
" SELECT eingang zeitpunkt,',katg,' FROM labor2a WHERE pat_id=',pid,'" & vbCrLf & _
" UNION " & vbCrLf & _
" SELECT zeitpunkt,',katg,' FROM labor1a l1 WHERE pat_id=',pid,' ORDER BY ',katg);" & vbCrLf & _
"PREPARE sc1 FROM @frag; " & vbCrLf & _
"execute sc1;" & vbCrLf & _
"DEALLOCATE PREPARE sc1; " & vbCrLf & _
"-- SELECT @frag; " & vbCrLf & _
"END;"
myEFrag (sql)

sql = "DROP PROCEDURE IF EXISTS `quelle`.`sensib`"
myEFrag sql
sql = _
"CREATE DEFINER=`praxis`@`%` PROCEDURE `sensib`(in pid INT(6), out obpath INT, out pz INT, out gz INT, out ergeb VARCHAR(1000)) " & vbCrLf & vbCrLf & _
"    READS SQL DATA " & vbCrLf & _
"    DETERMINISTIC " & vbCrLf & _
"    COMMENT 'zur Beurteilung der Sensibilität' " & vbCrLf & _
"BEGIN " & vbCrLf & _
"/*DECLARE done0 INT DEFAULT FALSE;*/ " & vbCrLf & _
"DECLARE pathzahl, geszahl, fundzahl INT; " & vbCrLf & _
"DECLARE mfr, mfl, vibir, vibil, vibgr, vibgl DECIMAL(2,1); " & vbCrLf & _
"DECLARE mfrs, mfls, vibirs, vibils, vibgrs, vibgls VARCHAR(200); " & vbCrLf & _
"DECLARE zp DATETIME; " & vbCrLf & _
"DECLARE cur0 cursOR fOR SELECT  anteil(mf_re,IF(u.art='usdm1' OR u.art='usdm2',5,3)), anteil(mf_li,IF(u.art='usdm1' OR u.art='usdm2',5,3)), anteil(vibr_ik_re,8), anteil(vibr_ik_li,8), anteil(vibr_gz_re,8), anteil(vibr_gz_li,8), " & vbCrLf & _
"    MF_re , MF_li, Vibr_IK_re, Vibr_IK_li, Vibr_GZ_re, Vibr_GZ_li, zeitpunkt " & vbCrLf & _
"  FROM usdm u WHERE pat_id = pid " & vbCrLf & _
"  AND u.zeitpunkt=(SELECT MAX(zeitpunkt) FROM usdm u1 WHERE u1.pat_id = u.pat_id); " & vbCrLf & _
"/*DECLARE CONTINUE HANDLER FOR NOT FOUND SET done0 = TRUE;*/ " & vbCrLf & _
"SET obpath = 1; " & vbCrLf & _
"SET pathzahl=0; " & vbCrLf & _
"SET geszahl=0; " & vbCrLf & _
"SET ergeb = ''; " & vbCrLf & _
"open cur0; " & vbCrLf & _
"SET fundzahl=found_rows(); " & vbCrLf & _
"IF fundzahl > 0 THEN " & vbCrLf & _
"read0: LOOP " & vbCrLf & _
"  FETCH cur0 INTO mfr, mfl, vibir, vibil, vibgr, vibgl, mfrs, mfls, vibirs, vibils, vibgrs, vibgls, zp; " & vbCrLf
sql = sql & _
"  IF Not (mfr = -1 AND mfl = -1 AND vibir = -1 AND vibil = -1 AND vibgr = -1 AND vibgl = -1) THEN " & vbCrLf & _
"    IF mfr<>-1 THEN IF mfr<0.8 THEN SET pathzahl=pathzahl+1; ELSE SET geszahl=geszahl+1; END IF; END IF; " & vbCrLf & _
"    IF mfl<>-1 THEN IF mfl<0.8 THEN SET pathzahl=pathzahl+1; ELSE SET geszahl=geszahl+1; END IF; END IF; " & vbCrLf & _
"    IF vibir <>-1 THEN IF vibir <0.6 THEN SET pathzahl=pathzahl+1; ELSE SET geszahl=geszahl+1; END IF; END IF; " & vbCrLf & _
"    IF vibil <>-1 THEN IF vibil <0.6 THEN SET pathzahl=pathzahl+1; ELSE SET geszahl=geszahl+1; END IF; END IF; " & vbCrLf & _
"    IF vibgr <>-1 THEN IF vibgr <0.6 THEN SET pathzahl=pathzahl+1; ELSE SET geszahl=geszahl+1; END IF; END IF; " & vbCrLf & _
"    IF vibgl <>-1 THEN IF vibgl <0.6 THEN SET pathzahl=pathzahl+1; ELSE SET geszahl=geszahl+1; END IF; END IF; " & vbCrLf & _
"    IF pathzahl > 1 OR (GesZahl < 4 AND pathzahl > 0) THEN " & vbCrLf & _
"     SET obpath=2; " & vbCrLf & _
"    ELSE " & vbCrLf & _
"     SET obpath=0; " & vbCrLf & _
"    END IF; " & vbCrLf & _
"    SET pz=pathzahl; " & vbCrLf & _
"    SET gz=geszahl; " & vbCrLf & _
"    SET ergeb = CONCAT(IF(obpath=1,'nicht untersucht',IF(obpath=2,'auffällig','unauffällig')),IF(obpath=1,'',CONCAT(': Monofil: ',mfrs,', ',mfls,', Vibr: ',vibirs,', ',vibils,', ',vibgls,', ',vibgls,' (',DATE_FORMAT(zp, '%d.%m.%y'),')'))); " & vbCrLf & _
"  END IF; " & vbCrLf & _
"  /*IF done0 THEN */LEAVE read0;/* END IF;*/ " & vbCrLf & _
"END LOOP; " & vbCrLf & _
"END IF; " & vbCrLf & _
"close cur0; " & vbCrLf & _
"END;"
myEFrag (sql)
      
      
sql = "DROP FUNCTION IF EXISTS `quelle`.`rang`"
myEFrag (sql)
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION  `quelle`.`rang`(tha text) RETURNS INT(1) " & _
    "NO SQL" & vbCrLf & _
    "DETERMINISTIC " & vbCrLf & _
    "COMMENT 'rang der Therapiearten' " & _
"BEGIN " & vbCrLf & _
"    CASE LOWER(Tha) " & vbCrLf & _
"      WHEN 'diät'THEN RETURN 0; " & vbCrLf & _
"      WHEN 'diät?'THEN RETURN 0; " & vbCrLf & _
"      WHEN 'oad' THEN RETURN 1; " & vbCrLf & _
"      WHEN 'glp1' THEN RETURN 2; " & vbCrLf & _
"      WHEN 'komb' THEN RETURN 3; " & vbCrLf & _
"      WHEN 'glp1ins' THEN RETURN 4; " & vbCrLf & _
"      WHEN 'ct' THEN RETURN 5; " & vbCrLf & _
"      WHEN '(i)ct?' THEN RETURN 6; " & vbCrLf & _
"      WHEN 'i/ct?' THEN RETURN 7; " & vbCrLf & _
"      WHEN 'glp1ict' THEN RETURN 8; " & vbCrLf & _
"      WHEN 'ict' THEN RETURN 9; " & vbCrLf & _
"      WHEN 'csii' THEN RETURN 10; " & vbCrLf & _
"      ELSE RETURN 0; " & vbCrLf & _
"    END CASE; " & vbCrLf & _
"  END; "
myEFrag (sql)

sql = "DROP FUNCTION IF EXISTS `quelle`.`aktq`"
myEFrag (sql)
sql = "CREATE  DEFINER=`praxis`@`%` FUNCTION `quelle`.`aktq` () RETURNS VARCHAR(5) " & vbCrLf & _
    "NO SQL" & vbCrLf & _
    "COMMENT 'aktuell verarbeitetes Quartal' " & vbCrLf & _
"  BEGIN " & vbCrLf & _
"    DECLARE q VARCHAR(5); " & vbCrLf & _
"    SET q = CONCAT((month(NOW()-INTERVAL " & frist & " DAY) -1) div 3 + 1,YEAR(NOW() - INTERVAL " & frist & " DAY)); " & vbCrLf & _
"    RETURN q; " & vbCrLf & _
"END; "
myEFrag (sql)

sql = "DROP FUNCTION IF EXISTS `quelle`.`quartal`"
myEFrag (sql)
sql = "CREATE  DEFINER=`praxis`@`%` FUNCTION `quelle`.`quartal` (datum DATETIME) RETURNS VARCHAR(5) " & vbCrLf & _
    "NO SQL" & vbCrLf & _
    "DETERMINISTIC " & vbCrLf & _
    "COMMENT 'aktuell verarbeitetes Quartal' " & vbCrLf & _
"  BEGIN " & vbCrLf & _
"    DECLARE q VARCHAR(5); " & vbCrLf & _
"    SET q = CONCAT((month(datum) -1) div 3 + 1,YEAR(datum)); " & vbCrLf & _
"   RETURN q; " & vbCrLf & _
"  END; "
myEFrag (sql)

sql = "DROP FUNCTION IF EXISTS `quelle`.`maxtha`"
myEFrag (sql)
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION  `quelle`.`maxtha`(pid text) RETURNS text " & _
    "READS SQL DATA " & _
    "COMMENT 'maxtha eines Quartals zu einem Patienten' " & _
"BEGIN " & vbCrLf & _
" DECLARE thera TYPE of therarten.therart DEFAULT ''; DECLARE akttha TYPE of therarten.therart; DECLARE aktzp DATETIME; DECLARE aktrang INT; DECLARE maxrang INT DEFAULT -1; " & vbCrLf & _
" DECLARE done INT DEFAULT FALSE; DECLARE csr cursOR fOR " & vbCrLf & _
"   SELECT therart,zp FROM therarten WHERE pat_id = pid AND zp<=qend() ORDER BY zp DESC, mpnr DESC; " & vbCrLf & _
" DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = TRUE; " & vbCrLf & _
" OPEN csr; " & vbCrLf & _
" read_loop:  LOOP " & vbCrLf & _
"  FETCH csr INTO akttha, aktzp; " & vbCrLf & _
"  IF done THEN LEAVE read_loop; END IF; " & vbCrLf & _
"   SET aktrang=rang(akttha); " & vbCrLf & _
"   IF aktrang > maxrang THEN " & vbCrLf & _
"    SET maxrang=aktrang; " & vbCrLf & _
"    SET thera=akttha; " & vbCrLf & _
"   END IF; " & vbCrLf & _
"  IF aktzp<qanf() THEN SET done = TRUE; END IF; " & vbCrLf & _
" END LOOP; " & vbCrLf & _
" CLOSE csr; " & vbCrLf & _
"  SELECT pat_id INTO akttha FROM namen LIMIT 1; " & vbCrLf & _
" RETURN thera; END; "
myEFrag (sql)


sql = "DROP FUNCTION IF EXISTS `qbeg`"
myEFrag (sql)
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `qbeg`( datum DATETIME ) RETURNS DATE " & _
"    DETERMINISTIC " & _
"    NO SQL" & _
"    COMMENT 'Quartalsanfang zu bestimmtem datum' " & _
"BEGIN" & vbCrLf & _
" DECLARE erg DATE; " & vbCrLf & _
" IF ISNULL(datum) THEN SET erg = 99990101; ELSE " & _
" SET erg = DATE(CONCAT(YEAR(datum),'-',(QUARTER(datum)-1)*3+1,'-01')); END IF; " & vbCrLf & _
" RETURN erg; " & vbCrLf & _
"END"
myEFrag (sql)

sql = "DROP FUNCTION IF EXISTS `vorquart`"
myEFrag (sql)
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `vorquart`(quart CHAR(5),zahl INTEGER) RETURNS CHAR(5) " & _
"    NO SQL" & _
"    DETERMINISTIC " & _
"    COMMENT 'zahl.tes Vorquartal zu quart' " & _
"BEGIN " & vbCrLf & _
" DECLARE erg CHAR(5); " & vbCrLf & _
" DECLARE qnr CHAR(1); " & vbCrLf & _
" DECLARE jnr CHAR(4); " & vbCrLf & _
" SET qnr=MID(quart,1,1); " & vbCrLf & _
" SET jnr=MID(quart,2,4); " & vbCrLf & _
" WHILE zahl>0 do " & vbCrLf & _
"  SET qnr=qnr-1; " & vbCrLf & _
"  IF qnr = 0 THEN " & vbCrLf & _
"   SET qnr=4; " & vbCrLf & _
"   SET jnr=jnr-1; " & vbCrLf & _
"  END IF; " & vbCrLf & _
"  SET zahl=zahl-1; " & vbCrLf & _
" END WHILE; " & vbCrLf & _
" SET erg=CONCAT(qnr,jnr); " & vbCrLf & _
" RETURN erg; " & vbCrLf & _
"END"
myEFrag (sql)
' sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `vorquart`(quart CHAR(5),zahl INTEGER) RETURNS CHAR(5) CHARSET utf8mb4 COLLATE utf8mb4_german2_ci " & _


sql = "DROP FUNCTION IF EXISTS `qbegs`"
myEFrag (sql)
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `qbegs`( quartal CHAR(5) ) RETURNS DATE " & _
"    NO SQL" & vbCrLf & _
"    DETERMINISTIC " & vbCrLf & _
"    COMMENT 'Quartalsanfang zu Quartal' " & _
"BEGIN " & vbCrLf & _
" DECLARE m CHAR(2); " & vbCrLf & _
" DECLARE q CHAR(1); " & vbCrLf & _
" DECLARE erg DATE; " & vbCrLf & _
" SET q = SUBSTR(quartal,1,1); " & vbCrLf & _
" CASE q WHEN '1' THEN SET m='01'; WHEN '2' THEN SET m='04'; WHEN '3' THEN SET m='07'; WHEN '4' THEN SET m='10'; END CASE; " & vbCrLf & _
" SET erg=DATE(CONCAT(SUBSTR(quartal,2),m,'01')); " & vbCrLf & _
" RETURN erg; " & vbCrLf & _
"END"
myEFrag (sql)

sql = "DROP FUNCTION IF EXISTS `qends`"
myEFrag (sql)
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `qends`( quartal CHAR(5) ) RETURNS DATETIME " & _
"    NO SQL" & vbCrLf & _
"    DETERMINISTIC " & vbCrLf & _
"    COMMENT 'Quartalsanfang zu Quartal' " & vbCrLf & _
"BEGIN " & vbCrLf & _
" DECLARE m CHAR(2); " & vbCrLf & _
" DECLARE q CHAR(1); " & vbCrLf & _
" DECLARE y CHAR(4); " & vbCrLf & _
" DECLARE erg DATETIME; " & vbCrLf & _
" SET q = SUBSTR(quartal,1,1); " & vbCrLf & _
" SET y = SUBSTR(quartal,2,4); " & vbCrLf & _
" case q when '1' THEN SET m='04'; when '2' THEN SET m='07'; when '3' THEN SET m='10'; when '4' THEN SET m='01'; SET y=y+1; END case; " & vbCrLf & _
" SET erg=DATE(CONCAT(y,m,'01'))-INTERVAL 1 second; " & vbCrLf & _
" RETURN erg; " & vbCrLf & _
"END"
myEFrag (sql)

sql = "DROP FUNCTION IF EXISTS `wia`;"
myEFrag sql
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `wia`( pid INT(6) UNSIGNED ) RETURNS VARCHAR(2) CHARSET utf8" & vbCrLf & _
    "    READS SQL DATA " & vbCrLf & _
    "    COMMENT 'wahrscheinlicher interner Arzt'" & vbCrLf & _
    "BEGIN" & vbCrLf & _
    "  DECLARE gsz, tkz, ddiff INT;" & vbCrLf & _
    "  DECLARE lgs, ltk DATE;" & vbCrLf & _
    "  DECLARE curgs cursOR fOR SELECT COUNT(0), MAX(zeitpunkt) FROM eintraege e WHERE pat_id = pid AND (e.art='gs' OR e.inhalt LIKE '%: gs' OR e.inhalt LIKE '%(gs)%');" & vbCrLf & _
    "  DECLARE curtk cursOR fOR SELECT COUNT(0), MAX(zeitpunkt) FROM eintraege e WHERE pat_id = pid AND (e.art='tk' OR e.inhalt LIKE '%: tk' OR e.inhalt LIKE '%(tk)%');" & vbCrLf & _
    "  OPEN curgs; FETCH curgs INTO gsz, lgs; CLOSE curgs;" & vbCrLf & _
    "  OPEN curtk; FETCH curtk INTO tkz, ltk; CLOSE curtk;" & vbCrLf & _
    "  IF gsz=0 AND tkz>0 THEN RETURN 'tk'; END IF;" & vbCrLf & _
    "  IF gsz>0 AND tkz=0 THEN RETURN 'gs'; END IF;" & vbCrLf & _
    "  IF gsz=0 AND tkz=0 THEN RETURN '-'; END IF;" & vbCrLf & _
    "  SET ddiff = DATEdiff(lgs,ltk);" & vbCrLf & _
    "  IF ddiff > 90 THEN RETURN 'gs'; END IF;" & vbCrLf & _
    "  IF ddiff < -90 THEN RETURN 'tk'; END IF;" & vbCrLf & _
    "  IF gsz> 4 * tkz THEN RETURN 'gs'; END IF;" & vbCrLf & _
    "  IF tkz> 4 * gsz THEN RETURN 'tk'; END IF;" & vbCrLf & _
    "  RETURN '?';" & vbCrLf & _
    "End"
myEFrag (sql)

sql = "DROP FUNCTION IF EXISTS `lHbA`;"
myEFrag (sql)
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `lHbA`( pid INT(6) UNSIGNED ) RETURNS DECIMAL(9,2) " & vbCrLf & _
"    READS SQL DATA " & vbCrLf & _
"    COMMENT 'letzter HbA1c' " & vbCrLf & _
"BEGIN " & vbCrLf & _
"  DECLARE zp1, zp2 DATE; " & vbCrLf & _
"  DECLARE Hb1, Hb2 VARCHAR(10); " & vbCrLf & _
"  DECLARE erg DECIMAL(9,2); " & vbCrLf & _
"  DECLARE done INT DEFAULT FALSE; " & vbCrLf & _
"  DECLARE cur1 cursOR fOR SELECT zeitpunkt, wert FROM laborneu WHERE pat_id = pid AND abkü RLIKE '^hba[1c]' AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM laborneu WHERE pat_id = pid AND abkü RLIKE '^hba[1c]'); " & vbCrLf & _
"  DECLARE cur2 cursOR fOR SELECT eingang zeitpunkt, wert FROM laborywert w LEFT JOIN laboryus u ON u.id=w.usid WHERE pat_id = pid AND abkü RLIKE '^hba[1c]' AND eingang = (SELECT MAX(eingang) FROM laborywert w LEFT JOIN laboryus u ON u.id=w.usid WHERE pat_id = pid AND abkü RLIKE '^hba[1c]'); " & vbCrLf & _
"  DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = TRUE; " & vbCrLf & _
"  OPEN cur1; fetch cur1 INTO zp1, Hb1; CLOSE cur1; " & vbCrLf & _
"  OPEN cur2; fetch cur2 INTO zp2, Hb2; CLOSE cur2; " & vbCrLf & _
"  IF NOT ISNULL(zp2) THEN IF zp2 > zp1 OR ISNULL(zp1) THEN SET zp1=zp2; SET Hb1=Hb2; END IF; END IF; " & vbCrLf & _
"  IF NOT ISNULL(zp1) THEN " & vbCrLf & _
"   SET erg=CONVERT(REPLACE(CONCAT('0',trim(Hb1)),',','.'),DECIMAL(9,2)); " & vbCrLf & _
"   RETURN erg; " & vbCrLf & _
"  END IF; " & vbCrLf & _
"  RETURN 0; " & vbCrLf & _
"End"
myEFrag (sql)

sql = "DROP FUNCTION IF EXISTS `lCre`;"
myEFrag (sql)
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `lCre`( pid INT(6) UNSIGNED ) RETURNS DECIMAL(9,2) " & vbCrLf & _
"    READS SQL DATA " & vbCrLf & _
"    COMMENT 'letztes Kreatinin' " & vbCrLf & _
"BEGIN " & vbCrLf & _
"  DECLARE zp1, zp2 DATE; " & vbCrLf & _
"  DECLARE Hb1, Hb2 VARCHAR(10); " & vbCrLf & _
"  DECLARE erg DECIMAL(9,2); " & vbCrLf & _
"  DECLARE done INT DEFAULT FALSE; " & vbCrLf & _
"  DECLARE cur1 cursOR fOR SELECT zeitpunkt, wert FROM laborneu WHERE pat_id = pid AND abkü IN ('CREAT','KRE02','KREA','KREA02','KRES') AND einheit = 'mg/dl' AND zeitpunkt = (SELECT MAX(zeitpunkt) FROM laborneu WHERE pat_id = pid AND abkü IN ('CREAT','KRE02','KREA','KREA02','KRES') AND einheit = 'mg/dl'); " & vbCrLf & _
"  DECLARE cur2 cursOR fOR SELECT eingang zeitpunkt, wert FROM laborywert w LEFT JOIN laboryus u ON u.id=w.usid WHERE pat_id = pid AND abkü IN ('CREAT','KRE02','KREA','KREA02','KRES') AND einheit = 'mg/dl' AND eingang = (SELECT MAX(eingang) FROM laborywert w LEFT JOIN laboryus u ON u.id=w.usid WHERE pat_id = pid AND abkü IN ('CREAT','KRE02','KREA','KREA02','KRES') AND einheit = 'mg/dl'); " & vbCrLf & _
"  DECLARE CONTINUE HANDLER FOR NOT FOUND SET done = TRUE; " & vbCrLf & _
"  OPEN cur1; fetch cur1 INTO zp1, Hb1; CLOSE cur1; " & vbCrLf & _
"  OPEN cur2; fetch cur2 INTO zp2, Hb2; CLOSE cur2; " & vbCrLf & _
"  IF NOT ISNULL(zp2) THEN IF zp2 > zp1 OR ISNULL(zp1) THEN SET zp1=zp2; SET Hb1=Hb2; END IF; END IF; " & vbCrLf & _
"  IF NOT ISNULL(zp1) THEN " & vbCrLf & _
"   SET erg=CONVERT(REPLACE(CONCAT('0',trim(Hb1)),',','.'),DECIMAL(9,2)); " & vbCrLf & _
"   RETURN erg; " & vbCrLf & _
"  END IF; " & vbCrLf & _
"  RETURN 0; " & vbCrLf & _
"End"
myEFrag (sql)

sql = "DROP PROCEDURE IF EXISTS `ltWert`;"
myEFrag (sql)
sql = "CREATE DEFINER=`praxis`@`%` PROCEDURE `ltWert`(in pid INT(6) UNSIGNED, abk VARCHAR(1000), einh VARCHAR(10)) " & vbCrLf & _
"    READS SQL DATA " & vbCrLf & _
"    COMMENT 'letztes Kreatinin' " & vbCrLf & _
"BEGIN " & vbCrLf & _
"  DECLARE ergeb DECIMAL(9,2) DEFAULT 0; " & vbCrLf & _
"  SET @sql1 = CONCAT('SELECT zeitpunkt, wert INTO @zp1, @w1 FROM laborneu WHERE pat_id = ',pid,' AND abkü IN (\'', abk, '\') AND einheit LIKE \'', einh,'%\' ORDER BY zeitpunkt DESC LIMIT 1'); " & vbCrLf & _
"  SET @sql2 = CONCAT('SELECT eingang zeitpunkt, wert INTO @zp2, @w2 FROM laborywert w LEFT JOIN laboryus u ON u.id=w.usid WHERE pat_id = ',pid,' AND abkü IN (\'', abk, '\') AND einheit LIKE \'', einh,'%\' ORDER BY eingang DESC LIMIT 1'); " & vbCrLf & _
"  PREPARE sc1 FROM @sql1; EXECUTE sc1; DEALLOCATE PREPARE sc1; " & vbCrLf & _
"  PREPARE sc2 FROM @sql2; EXECUTE sc2; DEALLOCATE PREPARE sc2; " & vbCrLf & _
"  IF NOT ISNULL(@zp2) THEN IF @zp2 > @zp1 OR ISNULL(@zp1) THEN SET @zp1=@zp2; SET @w1=@w2; END IF; END IF; " & vbCrLf & _
"  IF NOT ISNULL(@zp1) THEN " & vbCrLf & _
"   SET ergeb=zuBruch(@w1); " & vbCrLf & _
"  END IF; " & vbCrLf & _
"  SELECT ergeb ltWert, @zp1 Zp; " & vbCrLf & _
"END"
myEFrag (sql)

' wird noch benötigt für das Fotoprogramm, und epikrise
sql = "DROP FUNCTION IF EXISTS `dmtyp`;"
myEFrag sql
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `dmtyp`( pid INT(6) UNSIGNED ) RETURNS VARCHAR(1) CHARSET utf8" & vbCrLf & _
"    READS SQL DATA" & vbCrLf & _
"    COMMENT 'Diabetestyp (`1`,`2`,`s`,`u`,`g`)'" & vbCrLf & _
"BEGIN" & vbCrLf & _
"  DECLARE typ VARCHAR(1);" & vbCrLf & _
"  SET typ = (SELECT CASE SUBSTR(gicd,1,1) WHEN 'E' THEN CASE SUBSTR(gicd,3,1) WHEN 1 THEN '2' WHEN 0 THEN '1' WHEN 3 THEN 's' ELSE 'u' END WHEN 'O' THEN 'g' WHEN 'R' THEN 'p' ELSE '-' END FROM " & vbCrLf & _
"  (SELECT MIN(icd) gicd FROM diagview d WHERE pat_id = pid AND (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit in ('G',' '))) " & vbCrLf & _
"  ) i );" & vbCrLf & _
"RETURN typ;" & vbCrLf & _
"END"
' wenn z.B. erst E10.91 und dann O24.0
'"   AND d.diagdatum = (SELECT MAX(diagdatum) FROM diagnosen di WHERE di.pat_id = d.pat_id AND ((di.icd REGEXP '^E1[0-4]\.' ) OR (di.icd = 'O24.4' )) AND di.diagsicherheit NOT IN ('A','Z') AND COALESCE(di.Dggel,0)=0)" & vbCrLf
 myEFrag (sql)

sql = "DROP FUNCTION IF EXISTS `dmtypicd`;"
myEFrag sql
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `dmtypicd`(icd VARCHAR(6)) RETURNS VARCHAR(1) CHARSET utf8" & vbCrLf & _
"  NO SQL DETERMINISTIC" & vbCrLf & _
"    COMMENT 'Diabetestyp (`1`,`2`,`s`,`u`,`g`) nach ICD'" & vbCrLf & _
"BEGIN" & vbCrLf & _
"RETURN CASE SUBSTR(icd,1,1) WHEN 'E' THEN CASE SUBSTR(icd,3,1) WHEN 1 THEN '2' WHEN 0 THEN '1' WHEN 3 THEN 's' ELSE 'u' END WHEN 'O' THEN 'g' WHEN 'R' THEN 'p' ELSE '-' END ;" & vbCrLf & _
"END"
' wenn z.B. erst E10.91 und dann O24.0
'"   AND d.diagdatum = (SELECT MAX(diagdatum) FROM diagnosen di WHERE di.pat_id = d.pat_id AND ((di.icd REGEXP '^E1[0-4]\.' ) OR (di.icd = 'O24.4' )) AND di.diagsicherheit NOT IN ('A','Z') AND COALESCE(di.Dggel,0)=0)" & vbCrLf
myEFrag (sql)

' würde unter diesem Namen auch in pznbdt aufgerufen
sql = "DROP PROCEDURE IF EXISTS `fuellThaP`;"
myEFrag (sql)
sql = "CREATE DEFINER=`praxis`@`%` PROCEDURE `quelle`.`fuellThaP`(IN inpid TEXT) " & vbCrLf & _
"    MODIFIES SQL DATA " & vbCrLf & _
"    COMMENT 'fuellt die Tabelle tharten, pid 0 => alle; Algorithmus definiert in therinit, seit 11.12.20' " & vbCrLf & _
"proc: BEGIN " & vbCrLf & _
"-- RESET QUERY CACHE;" & vbCrLf & _
" DECLARE sqlt TEXT; " & vbCrLf
 Dim iru&
' Lese.ProgStart
 Call therinit
 sql = sql & " SET inpid=REPLACE(REPLACE(inpid,'""',''),'''','');" & vbCrLf & _
             " IF inpid IN ('','0') THEN SET inpid='0'; TRUNCATE therarten; END IF;" & vbCrLf
'             " IF inpid IN ('','0') THEN SET inpid='SELECT DISTINCT pat_id FROM namen'; END IF;" & vbCrLf
' sql = sql & "-- CASE WHEN LEFT(inpid,1)='''' THEN DO 0; ELSE SET inpid=CONCAT('''',inpid); END CASE;" & vbCrLf
' sql = sql & "-- CASE WHEN RIGHT(inpid,1)='''' THEN DO 0; ELSE SET inpid=CONCAT(inpid,''''); END CASE;" & vbCrLf
 For iru = 1 To 5
  Select Case iru
   Case 1, 3, 5
    sql = sql & " SET sqlt = CONCAT('" & REPLACE$(REPLACE$(REPLACE$(psql(iru), "'", "''"), "inpid=", "''',inpid,'''="), "(inpid)", "(',inpid,')") & "');" & vbCrLf
'   Case 1
'    sql = sql & " SET sqlt = CONCAT('" & replace$(replace$(replace$(psql(iru), "'", "''"), "FIND_IN_SET(pat_id,inpid)>0", "pat_id IN (',inpid,')"), " inpid ", " ''',inpid,''' ") & "');" & vbCrLf
''    sql = REPLACE(replace$(psql(iru), "FIND_IN_SET(pat_id,inpid)>0", "pat_id IN (" & pids & ")"), "inpid", pids)
'   Case 3, 5
'    sql = sql & " SET sqlt = CONCAT('" & replace$(replace$(psql(iru), "'", "''"), "inpid IN('''',''0'') OR FIND_IN_SET(x.pat_id,inpid)>0", " ''',inpid,''' IN ('''',''0'') OR x.pat_id IN (', inpid ,')") & "');" & vbCrLf
   Case Else
    sql = sql & " " & REPLACE$(psql(iru), "'", "''") & vbCrLf
  End Select
 Select Case iru
  Case 1, 3, 5
   sql = sql & _
   " PREPARE stmt FROM sqlt;" & vbCrLf & _
   " EXECUTE stmt;" & vbCrLf
 End Select ' case iru
 Select Case iru
  Case 2, 4, 5
   sql = sql & _
   " DEALLOCATE PREPARE stmt;" & vbCrLf
 End Select ' case iru
Next iru
 
#If alt Then
sql = sql & _
"SET SESSION GROUP_CONCAT_MAX_LEN=15000;" & vbCrLf & _
"-- SET @inpid=inpid;" & vbCrLf
sql = sql & _
"IF inpid=0 THEN DELETE FROM therarten; ELSE DELETE FROM therarten WHERE pat_id=inpid; END IF;" & vbCrLf & _
"SET @vzahl = ROW_COUNT(); " & vbCrLf & _
"INSERT INTO therarten(pat_id,zp,mpnr,therart,insart,grund,abspos,aktzeit,stbyte) " & vbCrLf & _
"SELECT pid,zp,mpnr,thart,ia,gru,abspos,NOW(),stbyte FROM ( " & vbCrLf & _
"WITH dsort AS ( -- da zum Vergleich mit dem Vorbefund zweimal zu verwenden " & vbCrLf & _
"SELECT RANK() OVER (PARTITION BY pid ORDER BY zp,MPNr) rang, i.* FROM ( " & vbCrLf & _
"  -- 1) Anamnese:" & vbCrLf & _
"  SELECT a.pat_id pid, -1 MPNr, aufndat Zp, aufndat Bis, 'CSII' Thart, 'Anamnese: Insulinpumpe' Gru,  0 ia, n.absPos, n.StByte, 0 FeldNr FROM anamnesebogen a LEFT JOIN namen n USING (pat_id) " & vbCrLf & _
"  WHERE insulinpumpe<>0 AND (inpid=0 OR a.pat_id=inpid) " & vbCrLf & _
" UNION -- 2) Rezepte für Pumpenzubehör, könnten für ein Jahr Pumpentherapie versprechen: " & vbCrLf & _
"  SELECT Pat_id pid,-2 MPNr, zeitpunkt Zp,ADDDATE(zeitpunkt,365) bis,'CSII' Thart, feldinh Gru,f.foid ia,FID absPos,Form_id StByte, feldnr FROM formular f " & vbCrLf & _
"  WHERE form_abk IN ('rp','lar','prp','plar') AND feld IN ('medikament','txtMedKey','VerordnungsZeile') AND feldinh RLIKE 'reservoir|rapid d link|rap d li|rapid-d li|tenderl|sure t|paradigm|veo|animas|cartridge|t-slim|t:slim|variosoft|trusteel|autosoft|ypsopump|insigh|omnipod' AND NOT feldinh LIKE '%menveo%'" & vbCrLf & _
"  AND (inpid=0 OR pat_id=inpid) " & vbCrLf & _
" UNION -- 3) Insulinpläne " & vbCrLf & _
" SELECT pat_id pid,-3 MPNr,qdm zp,qdm bis,'ICT' Thart, MID(NAME,p) Gru,-2 ia,b.absPos,b.StByte,0 FROM (SELECT IF(p1>p2,p1,p2) p, b.* FROM (SELECT INSTR(b.name,'insulin') p1, INSTR(b.name,'spritz') p2, b.* FROM briefe b) b) b WHERE b.name RLIKE '(insulin|spritz).*(plan|schema|tabelle)' " & vbCrLf & _
"  AND (inpid=0 OR pat_id=inpid) " & vbCrLf & _
" UNION -- 4) Medikamentenplan " & vbCrLf & _
"  SELECT Pid,MPNr,Zp,Zp bis,Thart,gru,ia,abspos,stbyte,feldnr FROM ( " & vbCrLf & _
"   SELECT Pid,MPNr,Zp,Zp bis " & vbCrLf
sql = sql & _
"       ,CASE WHEN SUM(pu) THEN 'CSII' " & vbCrLf & _
"       WHEN SUM(obmzi) OR sum(iz)>=3 THEN -- wenn Mahlzeiteninsulin oder mind. 3 x Insulin/d " & vbCrLf & _
"       CASE " & vbCrLf & _
"        WHEN SUM(glp) THEN 'GLP1ICT' " & vbCrLf & _
"        ELSE 'ICT' " & vbCrLf & _
"       END " & vbCrLf & _
"      WHEN SUM(iz)=2 THEN " & vbCrLf & _
"       CASE " & vbCrLf & _
"        WHEN SUM(glp) THEN 'GLP1Ins' " & vbCrLf & _
"        ELSE 'CT' " & vbCrLf & _
"       END " & vbCrLf & _
"      WHEN SUM(iz)=1 THEN -- Insulinzahl pro Tag " & vbCrLf & _
"       CASE " & vbCrLf & _
"        WHEN SUM(glp) THEN 'GLP1Ins' " & vbCrLf & _
"        WHEN SUM(oboad) THEN 'Komb' " & vbCrLf & _
"        ELSE 'CT' " & vbCrLf & _
"       END " & vbCrLf & _
"      ELSE -- WHEN SUM(iz)=0 THEN " & vbCrLf & _
"       CASE " & vbCrLf & _
"        WHEN SUM(glp) THEN 'GLP1' " & vbCrLf & _
"        WHEN SUM(oboad) THEN 'OAD' " & vbCrLf & _
"        ELSE 'Diät' " & vbCrLf & _
"       END " & vbCrLf
sql = sql & _
"     END Thart " & vbCrLf & _
"   ,GROUP_CONCAT(DISTINCT gru SEPARATOR '') gru,MAX(ia) ia,abspos,stbyte,MIN(feldnr) feldnr -- Grund, Insulinart " & vbCrLf & _
"   FROM ( " & vbCrLf & _
"    SELECT Pid,MPNr,Zp,Med,Pu,oboad,IF(obin,ezm,0) iz,ezm AND ia=1 obmzi " & vbCrLf & _
"     ,IF(eztm,glp,0) glp,IF(ezm,obin,0) ins,IF(ezm,ia,0) ia " & vbCrLf & _
"     ,IF(pu||if(eztm,glp,0)||oboad||(ezm&&ia=1)||if(obin,ezm,0),CONCAT('/',Med,' '),'') gru " & vbCrLf & _
"     ,IF(pu||if(eztm,glp,0)||oboad||(ezm&&ia=1)||if(obin,ezm,0),Feldnr,NULL) FeldNr,absPos,StByte " & vbCrLf & _
"    FROM ( " & vbCrLf & _
"     SELECT pid,MPNr,zp,Med,pu,ez,wglp,ohneE,IF(ez,oad,0)oboad,glp,obin,ia,FeldNr,absPos,StByte " & vbCrLf & _
"      ,IF(ez,ez,ohnee) ezm -- Eintragszahl modifziert " & vbCrLf & _
"      ,IF(ez,ez,wglp&&ohneE) eztm -- Eintragszahl teilmodifiziert " & vbCrLf & _
"     FROM ( " & vbCrLf & _
"      SELECT mp.Pat_id pid,mp.MPNr MPNr,mp.Zeitpunkt Zp,mp.Medikament Med,ma.puzu<>0 pu " & vbCrLf & _
"       ,MAX((COALESCE(mp.mo,'')<>'')+(COALESCE(mp.mi,'')<>'')+(COALESCE(mp.nm,'')<>'')+(COALESCE(mp.ab,'')<>'')+(COALESCE(mp.zn,'')<>'')+if(glp1<>0 AND mp.Medanfang RLIKE 'OZEMPIC|TRULICITY',1,0)) ez " & vbCrLf & _
"       ,glp1<>0 AND mp.MedAnfang RLIKE 'OZEMPIC|TRULICITY' wglp -- Wochen-GLP-1 " & vbCrLf & _
"       ,pzn<>0 AND concat(mp.Bemerkung,' ',mp.Grund) NOT RLIKE 'Pause|abgesetzt|beendet|zur Zeit nicht' ohneE -- ohne Eintrag im neuen Medplan, aber vermutlich angewandt " & vbCrLf & _
"       ,ma.glib<>0 OR ma.metf<>0 OR ma.gluci<>0 OR ma.shglin<>0 OR ma.dpp4<>0 OR ma.sglt2<>0 OR ma.sonstad<>0 oad " & vbCrLf & _
"       ,glp1<>0 glp " & vbCrLf & _
"       ,ma.ins<>0 OR ma.anal<>0 obin -- ob Insulin " & vbCrLf & _
"       ,IF(insart='' OR ISNULL(insart),0,insart) ia -- Insulinart: 1= schnell, 2 = langsam, 3 = Misch " & vbCrLf
sql = sql & _
"       ,mp.FeldNr,mp.absPos,mp.StByte -- zur Zeit eher überflüssige Felder " & vbCrLf & _
"      FROM wmedplan mp LEFT JOIN medarten ma ON mp.medanfang= ma.medikament " & vbCrLf & _
"      WHERE (inpid=0 OR mp.pat_id=inpid) " & vbCrLf & _
"      GROUP BY mp.pat_id,mpnr,mp.zeitpunkt,ma.id -- z.B. versch.Toujeo-Zeilen für v. Zuckerwerte" & vbCrLf & _
"     ) i " & vbCrLf & _
"    ) i " & vbCrLf & _
"   ) i GROUP BY pid,MPNr,zp -- 13.12.21: gleiche MPNr für versch. Zeitpunkte! wohl durch HB_-Import" & vbCrLf & _
"  ) i " & vbCrLf & _
" ) i " & vbCrLf & _
") -- with ...; letzte Therapieart, letzter Zeitpunkt: " & vbCrLf & _
"SELECT d.*, COALESCE(LAG(thart,1) OVER (PARTITION BY pid ORDER BY zp,MPNr),'') lthart, COALESCE(LAG(zp,1) OVER (PARTITION BY pid ORDER BY zp,MPNr),'1900-01-01') lzp, COALESCE(LAG(ia,1) OVER (PARTITION BY pid ORDER BY zp,MPNr),'-10') lia " & vbCrLf & _
"FROM dsort d " & vbCrLf & _
" WHERE NOT EXISTS (SELECT 1 FROM dsort d1 WHERE d.pid=d1.pid AND d1.rang<d.rang AND (d.zp > d1.zp AND d.zp <= d1.bis)) -- Gültigkeitsende wirken lassen (bes. Pumpenrezept) " & vbCrLf & _
") i -- nur Wechsel anzeigen, nicht von CSII/ICT auf Diät, nicht im Karenzzeitraum (1a nach Pumpenrezept oder 92 Tage nach Insulinplan): " & vbCrLf & _
"WHERE lthart<>thart AND NOT (thart='Diät' AND lthart IN ('CSII','ICT','GLP1ICT') AND NOT EXISTS (SELECT 1 FROM sws WHERE pat_id=i.pid AND voret BETWEEN i.lzp AND i.zp)) " & vbCrLf & _
"AND NOT (lia=-2 AND ia=2 AND NOT (thart<>lthart AND thart IN ('GLP1','GLP1Ins','GLP1ICT')) AND zp BETWEEN lzp AND ADDDATE(lzp,92)) " & vbCrLf & _
"ORDER BY pid,zp,MPNr; " & vbCrLf & _
"SELECT @vzahl vzahl, ROW_COUNT() zahl; " & vbCrLf & _
"UPDATE anamnesebogen a SET therakt =(SELECT therart FROM therarten WHERE pat_id=a.pat_id ORDER BY zp DESC, mpnr DESC LIMIT 1) WHERE inpid=0 OR pat_id=inpid;"
#End If
sql = sql & "END"
myEFrag (sql)


sql = "DROP FUNCTION IF EXISTS `RRsyst`;"
myEFrag sql
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `RRsyst` (rr text) RETURNS text " & vbCrLf & _
"    NO SQL" & vbCrLf & _
"    DETERMINISTIC " & vbCrLf & _
"BEGIN " & vbCrLf & _
"DECLARE pos,ppu,schonkomma INT default 0;" & vbCrLf & _
"DECLARE ch VARCHAR(1) default '';" & vbCrLf & _
"DECLARE rueck text default '';" & vbCrLf & _
"SET pos=INSTR(rr,'/');" & vbCrLf & _
"IF pos = 0 THEN" & vbCrLf & _
" SET pos=INSTR(rr,' mm Hg');" & vbCrLf & _
" IF pos=0 THEN SET pos=INSTR(rr,' mmHg'); END IF;" & vbCrLf & _
"-- IF pos=0 THEN RETURN '0'; END IF;" & vbCrLf & _
" SET ppu=INSTR(rr,'letzten');" & vbCrLf & _
" IF ppu>0 AND INSTR(rr,'Messungen')<>0 THEN SET pos=ppu; END IF;" & vbCrLf & _
"ELSE" & vbCrLf & _
" SET ppu=INSTR(binary rr,'Puls');" & vbCrLf & _
" IF ppu>0 AND ppu<pos THEN SET pos=ppu; END IF;" & vbCrLf & _
"END IF;" & vbCrLf & _
"IF pos=0 THEN SET pos=LENGTH(rr)+1; END IF;" & vbCrLf & _
"markc:  Loop" & vbCrLf & _
" SET pos=pos-1;" & vbCrLf & _
" IF pos<=0 THEN IF rueck='' THEN SET rueck='0'; END IF; LEAVE markc; END IF;" & vbCrLf & _
" SET ch = SUBSTRING(rr,pos,1);" & vbCrLf & _
" IF ch REGEXP '[0-9,]' AND (schonkomma=0 OR ch<>',') THEN" & vbCrLf & _
"  IF ch<>',' OR rueck<>'' THEN SET rueck = CONCAT(ch,rueck); END IF;"
sql = sql & vbCrLf & _
"  IF ch=',' THEN SET schonkomma=1; END IF;" & vbCrLf & _
" ELSEIF (ch<>' ' AND rueck<>'') OR (ch=' ' AND rueck>50) THEN" & vbCrLf & _
"  LEAVE markc;" & vbCrLf & _
" END IF;" & vbCrLf & _
"END LOOP;" & vbCrLf & _
"SET pos=INSTR(rueck,',');" & vbCrLf & _
"IF pos=1 THEN" & vbCrLf & _
" SET rueck=MID(rueck,2);" & vbCrLf & _
"ELSEIF pos>1 THEN SET rueck=LEFT(rueck,pos-1);" & vbCrLf & _
"END IF;" & vbCrLf & _
"IF rueck<30 THEN SET rueck=0; END IF;" & vbCrLf & _
"IF CAST(rueck As Integer)>50000 THEN" & vbCrLf & _
" IF CAST(rueck As Integer)>1000000 THEN" & vbCrLf & _
"  SET pos=LENGTH(rueck);" & vbCrLf & _
"  IF MID(rueck,pos-3,1)='7' THEN" & vbCrLf & _
"   SET rueck=LEFT(rueck,3);" & vbCrLf & _
"  END IF;" & vbCrLf & _
" ELSE" & vbCrLf & _
"  SET pos=LENGTH(rueck);" & vbCrLf & _
"  IF MID(rueck,pos-2,1)='7' THEN" & vbCrLf & _
"   SET rueck=LEFT(rueck,pos-3);" & vbCrLf & _
" END IF;" & vbCrLf & _
" END IF;" & vbCrLf & _
"END IF;"
sql = sql & vbCrLf & _
"IF CAST(rueck AS INT)>300 THEN" & vbCrLf & _
" SET rueck=LEFT(rueck,3);" & vbCrLf & _
"END IF;" & vbCrLf & _
"RETURN rueck;" & vbCrLf & _
"END"
myEFrag (sql)

sql = "DROP FUNCTION IF EXISTS `RRdiast`;"
myEFrag sql
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `RRdiast` (rr text) RETURNS text " & vbCrLf & _
"    NO SQL" & vbCrLf & _
"    DETERMINISTIC " & vbCrLf & _
"BEGIN" & vbCrLf & _
"DECLARE pos,len INT;" & vbCrLf & _
"DECLARE ch VARCHAR(1);" & vbCrLf & _
"DECLARE rueck TEXT DEFAULT '';" & vbCrLf & _
"SET pos=INSTR(rr,'/');" & vbCrLf & _
"SET len=LENGTH(rr);" & vbCrLf & _
"IF pos=0 THEN RETURN '0'; END IF;" & vbCrLf & _
"markc:  Loop" & vbCrLf & _
" SET pos=pos+1;" & vbCrLf & _
" IF pos>len THEN LEAVE markc; END IF;" & vbCrLf & _
" SET ch = SUBSTRING(rr,pos,1);" & vbCrLf & _
" IF ch REGEXP '[0-9]' THEN SET rueck = CONCAT(rueck,ch); ELSEIF (ch<>' ' OR rueck<>'') THEN LEAVE markc; END IF;" & vbCrLf & _
"END LOOP;" & vbCrLf & _
"IF rueck='' THEN SET rueck=0; END IF;" & vbCrLf & _
"IF CAST(rueck AS INT)>199 THEN SET rueck=LEFT(rueck,2); END IF;" & vbCrLf & _
"RETURN rueck;" & vbCrLf & _
"END"
myEFrag (sql)

' s. holRRZahl()
sql = "DROP FUNCTION IF EXISTS `RRzahl`;"
myEFrag sql
' MID( und INSTR( achten nicht auf Klein- oder Großschreibung
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `RRzahl` (Bemgr text) RETURNS INTEGER " & vbCrLf & _
"    DETERMINISTIC " & vbCrLf & _
"    NO SQL" & vbCrLf & _
"BEGIN" & vbCrLf & _
"DECLARE p0, pos, ppu, len INT;" & vbCrLf & _
"DECLARE muster VARCHAR(13) DEFAULT '[0123456789]';" & vbCrLf & _
"DECLARE ch VARCHAR(1);" & vbCrLf & _
"DECLARE ges INT DEFAULT '0';" & vbCrLf & _
"SET pos=INSTR(Bemgr,'MITTELWERT'); " & vbCrLf & _
"IF pos=0 THEN SET pos=INSTR(Bemgr,'DURCHSCHNITT'); IF pos>0 THEN SET pos=pos+12; END IF; " & _
"ELSE SET pos=pos+10; END IF; " & _
"IF pos=0 THEN " & _
" RETURN 1; " & _
"ELSE "
sql = sql & vbCrLf & _
" SET p0=pos; " & vbCrLf & _
" SET len=LENGTH(Bemgr); " & _
" SET ppu=INSTR(Bemgr,'PULS'); " & vbCrLf & _
" IF ppu>0 AND ppu>p0 THEN SET len=ppu; END IF; " & vbCrLf & _
" wahr: WHILE pos<=len do" & vbCrLf & _
"  SET ch=MID(Bemgr,pos,1);" & vbCrLf & _
"  SET pos=pos+1;" & vbCrLf & _
"  IF ch=',' THEN LEAVE wahr; END IF;" & vbCrLf & _
"  IF INSTR(muster,ch) THEN SET ges=10*ges+ch; ELSE IF ges>0 THEN IF ch='-' THEN IF ges=24 THEN SET ges=1; END IF; LEAVE wahr; ELSE IF ch=' ' THEN IF INSTR(MID(Bemgr,pos),'Tage')=1 THEN SET ges=1; END IF; LEAVE wahr; ELSE IF ch='/' OR ch='.' THEN SET ges=1; LEAVE wahr; END IF; END IF; END IF; END IF; END IF;" & vbCrLf & _
"  IF ges>999 THEN SET ges=1; LEAVE wahr; END IF;" & vbCrLf & _
" END WHILE wahr;" & vbCrLf & _
"END IF;" & vbCrLf & _
"IF ges=0 THEN SET ges=1; END IF;" & vbCrLf & _
"RETURN ges;" & vbCrLf & _
"END;"
myEFrag (sql)
' ch REGEXP muster fast genauso schnell wie INSTR(muster,ch) (12,3 vs 11,3s)

#If problematisch Then
    'Laborview mit Parameter, aufzurufen über:
    ' "SELECT lab.* FROM (SELECT @patid:=14 nix) nul, geslaba lab;"
    sql = "DROP FUNCTION IF EXISTS `patid`"
    myEFrag sql
    sql = "CREATE DEFINER=`praxis`@`%` FUNCTION patid() RETURNS INTEGER READS SQL DATA RETURN @patid;"
    myEFrag sql
    sql = "DROP FUNCTION IF EXISTS `abkue`"
    myEFrag sql
    sql = "CREATE DEFINER=`praxis`@`%` FUNCTION abkue() RETURNS CHAR(255) READS SQL DATA RETURN @abkue;"
    myEFrag sql
    sql = "DROP FUNCTION IF EXISTS `einheit`"
    myEFrag sql
    sql = "CREATE DEFINER=`praxis`@`%` FUNCTION einheit() RETURNS CHAR(255) READS SQL DATA RETURN @einheit;"
    myEFrag sql
    sql = "DROP FUNCTION IF EXISTS `zpkl`"
    myEFrag sql
    sql = "CREATE DEFINER=`praxis`@`%` FUNCTION zpkl() RETURNS DATETIME READS SQL DATA RETURN CAST(@zpkl As DateTIME);"
    myEFrag sql
    sql = "DROP FUNCTION IF EXISTS `wertkl`"
    myEFrag sql
    sql = "CREATE DEFINER=`praxis`@`%` FUNCTION wertkl() RETURNS INTEGER READS SQL DATA RETURN @wertkl;"
    myEFrag sql
    
    sql = "DROP VIEW IF EXISTS `_geslab`;"
    myEFrag sql
    sql = "CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `_geslab` AS " & vbCrLf & _
    "SELECT * FROM labor2a WHERE pat_id = patid() AND (abkue()='' OR abkü RLIKE abkue()) AND (zpkl()=0 OR zeitpunkt<zpkl()) AND (wertkl()=0 OR wert<wertkl()) AND (einheit()='' OR einheit=einheit())" & vbCrLf & _
    " UNION " & vbCrLf & _
    "SELECT * FROM labor1a WHERE pat_id = patid() AND (abkue()='' OR abkü RLIKE abkue()) AND (zpkl()=0 OR zeitpunkt<zpkl()) AND (wertkl()=0 OR wert<wertkl()) AND (einheit()='' OR einheit=einheit());"
    myEFrag sql
    
    sql = "DROP VIEW IF EXISTS `geslab`;"
    myEFrag sql
    sql = "CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `geslab` AS " & vbCrLf & _
    "SELECT * FROM _geslab GROUP BY zeitpunkt,abkü,wert ORDER BY zeitpunkt DESC,abkü,wert,einheit,nb;"
    myEFrag sql
    
    ' Gesamtlabor für den Laufzettel, alle Werte des Synacthentests werden angezeigt
    sql = "DROP VIEW IF EXISTS `geslaba`;"
    myEFrag sql
    sql = "CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `geslaba` AS " & vbCrLf & _
    "SELECT GROUP_CONCAT(wert SEPARATOR '<br>') wert,abkü, einheit,Zeitpunkt,NB,gruppe,reihe,uNg,oNg,Pfad,Kommentar,Labor,Langtext,pat_id FROM " & vbCrLf & _
    " (SELECT * FROM " & vbCrLf & _
    "  (SELECT wert,abkü,einheit,zeitpunkt,nb,gruppe,reihe,uNg,oNg,Pfad,Kommentar,Labor,Langtext,pat_id FROM labor2a WHERE pat_id=patid() " & vbCrLf & _
    "    UNION " & vbCrLf & _
    "   SELECT wert,abkü,einheit,zeitpunkt,nb,gruppe,reihe,uNg,oNg,Pfad,Kommentar,Labor,Langtext,pat_id FROM labor1a l1 WHERE pat_id=patid()) i " & vbCrLf & _
    "  GROUP BY zeitpunkt,abkü,wert) i " & vbCrLf & _
    " GROUP BY abkü,einheit,nb,zeitpunkt" & vbCrLf & _
    "ORDER BY zeitpunkt DESC,abkü,wert,einheit,nb;"
    myEFrag sql
#End If ' problematisch
    
sql = "DROP VIEW IF EXISTS `dtypen`;"
myEFrag sql
sql = "CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `dtypen` AS " & vbCrLf & _
"SELECT Pat_id, icd gicd, ityp " & vbCrLf & _
", IF(obsek AND ityp IN ('1','2'),'s',ityp) ttyp, diagtext, obsek " & vbCrLf & _
"FROM ( " & vbCrLf & _
" SELECT n.pat_id, d.icd, d.diagtext, COALESCE(diagtext RLIKE 'pan[ck]r|sekund',0) obsek " & vbCrLf & _
" , CASE LEFT(icd,1) WHEN 'E' THEN CASE MID(icd,3,1) WHEN 1 THEN '2' WHEN 0 THEN '1' WHEN 3 THEN 's' ELSE 'u' END WHEN 'O' THEN 'g' WHEN 'R' THEN 'p' ELSE '-' END ityp " & vbCrLf & _
" , RANK() OVER(PARTITION BY n.pat_id ORDER BY obsek DESC,LEFT(d.gicd,3),MID(d.gicd,5,1)=9 DESC,MID(d.gicd,5) DESC,LENGTH(d.diagtext) DESC) rang " & vbCrLf & _
" FROM namen n " & vbCrLf & _
" LEFT JOIN diagview d ON d.pat_id=n.pat_id AND (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit in ('G',' '))) GROUP BY n.pat_id,d.gicd,d.diagtext " & vbCrLf & _
") i " & vbCrLf & _
"WHERE rang=1;"
myEFrag sql
    
sql = "DROP VIEW IF EXISTS `dtypview`;"
myEFrag sql
sql = "CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `dtypview` AS " & vbCrLf & _
"SELECT n.pat_id, d.dmtyp FROM namen n" & vbCrLf & _
"LEFT JOIN (SELECT pat_id,CASE WHEN icd RLIKE '^E' THEN CAST(MID(icd,3,1)+1 AS CHAR) WHEN icd RLIKE '^R' THEN 'p' ELSE 'g' END dmtyp " & vbCrLf & _
"FROM diagview d" & vbCrLf & _
"WHERE (d.gicd REGEXP '^E1[0-4]\.|^R73' OR (d.icd='O24.4' AND d.Dggel=0 AND d.diagsicherheit in ('G',' '))) GROUP BY pat_id) d" & vbCrLf & _
"ON d.pat_id=n.pat_id;"
myEFrag sql
    
    
' ausführliche Performanceanalyse am 12.12.20 durchgeführt => Ergebnis aber fehlerhaft, deshalb am 13.12.21 korrigiert
sql = "DROP VIEW IF EXISTS `wmedplan`;"
myEFrag sql
'sql = "CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `wmedplan` AS " & vbCrLf & _
"SELECT mp.* FROM medplan mp INNER JOIN " & vbCrLf & _
"(SELECT pat_id, MAX(zeitpunkt) zp, MAX(MPNr) mpnr FROM medplan imp GROUP BY imp.pat_id,imp.tag) i " & vbCrLf & _
"ON mp.pat_id=i.pat_id AND mp.zeitpunkt=i.zp AND mp.mpnr=i.mpnr;"
sql = "CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `wmedplan` AS " & vbCrLf & _
 "SELECT * FROM medplan mp " & vbCrLf & _
 "WHERE NOT EXISTS (SELECT 0 FROM medplan WHERE pat_id=mp.pat_id AND DATE(zeitpunkt)=DATE(mp.zeitpunkt) AND (zeitpunkt>mp.zeitpunkt OR (zeitpunkt=mp.ZeitPunkt AND mpnr>mp.mpnr)))"
myEFrag sql
    

sql = "DROP VIEW IF EXISTS `aktlue`;"
myEFrag sql
myEFrag "DROP TABLE IF EXISTS `aktlue`"
sql = "CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `aktlue` AS " & vbCrLf & _
"SELECT ido, nameo, vno, titelto, kvnro, ide, ue.* FROM (" & vbCrLf & _
"SELECT id ido,name nameo, vorname vno, titelt titelto, kvnr kvnro" & vbCrLf & _
", (SELECT MIN(id) FROM liuez e WHERE telefon = l.telefon AND fachgruppe NOT LIKE '% ehem%' AND l.fachgruppe LIKE '% ehem%') ide " & vbCrLf & _
"FROM liuez l) i " & vbCrLf & _
"LEFT JOIN liuez ue ON ue.id=if(ISNULL(ide),ido,ide)"
myEFrag sql

Const BZein$ = "'[0-9]{1,3}(,[0-9]?|)[ ]?([m]?mg(/dl|)[,]?|aktuell|)( nüchtern| [0-9]{1,2}(\\.[0-9]{2}|) Uhr[,]?|)[ ]?((\\(|)(mit |)|)Biosen(\\)|)|biosen (gemessen|verglichen)|BZ.*[0-9]{2,}.*mit Biosen|Kontrolle mit Biosen|Biosen.*EDTA|Biosen(Messung|messwert|sor|[ ]{0,2}m|_|s|se|meßwert| \\(plasma\\)| |)(([,:]?[ ]?|\\()[0-9]{1,2}([.:][0-9]{2}|)(\\)|[ ]?(Uhr[:]?|h nü)[ ]?|)|Messvergleich|: Kontrolle zum vorherigen Wert|)[>.;:]*[ ]*([ ]*BZ[:]?|)[ ]*(p[\.]?p |postprandial|nach dem Frühstück|n\.M |nach 1h |nach 2 be |nach [0-9]{1,2}[ ]?min(.|uten)( warten|)|hier |.*gegessen.*|[ ]*(\\(|)nü(\.[:]?|chtern(\\)|))|Leihgerät|re.|[0-9]{1,3} min.*|nach dem Abendessen |)[ ]*[)>.;:]*[ ]*[0-9]+[ ]?[.,/]?[ ]?[0-9o]{0,2}([ ,]{0,2}$|: |[ #/]*([=0-][ ]*[0-9]+(,0|)[ ]*|)([.,] | AB:|\\(|[;,] Ab|\\+ [0-9]|nochmal|verglichen|Kontrolle |spe|Abw.|\., Ab|,mg|m/|/dl|bz/mg|smg|m[, ]g|mm|mg|ng/dl|g/dl|mmg|mh|gm|m%|%|/min|mmHg|,g/|kg|mmmg| Abwei|[,]? umgerech|[0-9]{2}|-))'"
Const BZaus$ = "'Biosen(:|)( [-~] |[ ]*)mg/dl|biosen:$'"
Const BZausKontr$ = "'Biosen:$|Biosenglucose,|Antibiosen|Biosen: ~ [^ 0-9]|^([^0-9]+|)Biosen([^0-9]+|)$|Biosen:[ ]{1,2}(-|,|mg/dl)|Biosen[s]?[:]?[ ]+mg/dl|Biosen:[ ]?-(-|#|.|)[ ]+mg|OGTT mit Biosen|Schuhen, Biosen|gleich nochmal mit Biosen|als unser Biosen|Biosen schon abgeschaltet|Biosen(-Gerät|) (schon aus|am )|Biosen war|Biosen zu hoch|Biosen funktioniert|Biosen Geräte hat|Biosen aus( [^E]|$)|Biosen[ ]?(heute|derzeit|) nicht|biosen[^0-9]*(kaput|defekt)|morgen.*Biosen-Gerät|Biosen unauffällig|Biosen mißt.*falsch|Biosentic|Biosen tauschen'"

sql = "DROP VIEW IF EXISTS `BiosenMessung`;"
sql = "CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `BiosenMessung` AS " & vbCrLf & _
"SELECT pat_id,id,zeitpunkt,art,inhalt FROM eintraege WHERE (art IN ('bz','ogtt') OR (inhalt RLIKE 'biosen' AND" & vbCrLf & _
"inhalt RLIKE " & BZein & " AND NOT inhalt RLIKE " & BZaus & "))"
myEFrag sql

' wenn die beiden folgenden Abfragen beide keine Ergebnisse erzeugen, sind die Konstanten richtig
sql = "DROP VIEW IF EXISTS `BiosenDoppelt`;"
sql = "CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `BiosenDoppelt` AS " & vbCrLf & _
"SELECT pat_id,id,zeitpunkt,art,inhalt FROM eintraege WHERE inhalt RLIKE 'biosen'" & vbCrLf & _
"AND (inhalt RLIKE 'biosen' AND inhalt RLIKE " & BZein & " AND NOT inhalt RLIKE " & BZaus & ")" & vbCrLf & _
"AND inhalt RLIKE " & BZausKontr
myEFrag sql

sql = "DROP VIEW IF EXISTS `Biosengarnicht`;"
sql = "CREATE OR REPLACE ALGORITHM=UNDEFINED DEFINER=`praxis`@`%` SQL SECURITY DEFINER VIEW `Biosengarnicht` AS " & vbCrLf & _
"SELECT pat_id,id,zeitpunkt,art,inhalt FROM eintraege WHERE inhalt RLIKE 'biosen'" & vbCrLf & _
"AND NOT (inhalt RLIKE 'biosen' AND inhalt RLIKE " & BZein & " AND NOT inhalt RLIKE " & BZaus & ")" & vbCrLf & _
"AND NOT inhalt RLIKE " & BZausKontr
myEFrag sql

myEFrag sql

' Hilfsprogramm für Glucose 0, 1 und 2
sql = "DROP FUNCTION IF EXISTS `_vorNr`"
myEFrag sql
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `_vorNr`(inh VARCHAR(1000),p0 INTEGER) RETURNS INT(3) no sql " & vbCrLf & _
"BEGIN " & vbCrLf & _
"DECLARE muster VARCHAR(13) DEFAULT '[0123456789]';" & vbCrLf & _
"DECLARE ch VARCHAR(1);" & vbCrLf & _
"DECLARE ges VARCHAR(15) DEFAULT '';" & vbCrLf & _
"IF p0>0 THEN" & vbCrLf & _
" l1: WHILE p0>0 do" & vbCrLf & _
"  SET ch=MID(inh,p0,1);" & vbCrLf & _
"  SET p0=p0-1;" & vbCrLf & _
"  IF ch REGEXP muster THEN SET ges=ch; LEAVE l1; END IF;" & vbCrLf & _
" END WHILE l1;" & vbCrLf & _
" IF ges<>'' THEN" & vbCrLf & _
"  l2: WHILE p0>0 do" & vbCrLf & _
"   SET ch=MID(inh,p0,1);" & vbCrLf & _
"   IF ch REGEXP muster THEN" & vbCrLf & _
"    SET ges=CONCAT(ch,ges);" & vbCrLf & _
"   ELSEIF ch = '.' OR ch = ',' THEN" & vbCrLf & _
"    SET ges=CONCAT('.',ges);" & vbCrLf & _
"   ELSE" & vbCrLf & _
"    LEAVE l2;" & vbCrLf & _
"   END IF;" & vbCrLf & _
"   SET p0=p0-1;" & vbCrLf & _
"  END WHILE l2;"
sql = sql & vbCrLf & _
"  RETURN ROUND(ges);" & vbCrLf & _
" ELSE" & vbCrLf & _
"  RETURN 0;" & vbCrLf & _
" END IF;" & vbCrLf & _
" RETURN p0;" & vbCrLf & _
"END IF;" & vbCrLf & _
"RETURN 0;" & vbCrLf & _
"END"
myEFrag sql

'Glucose 0 im OGTT
sql = "DROP FUNCTION IF EXISTS `g0`"
myEFrag sql
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `g0`(inh VARCHAR(1000)) RETURNS INT(3)" & vbCrLf & _
"    NO sql " & vbCrLf & _
"BEGIN " & vbCrLf & _
"DECLARE p0 INTEGER DEFAULT INSTR(inh,'mg/dl');" & vbCrLf & _
"RETURN _vorNr(inh,p0);" & vbCrLf & _
"END"
myEFrag sql

'Glucose 1 im OGTT
sql = "DROP FUNCTION IF EXISTS `g1`"
myEFrag sql
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `g1`(inh VARCHAR(1000)) RETURNS INT(3)" & vbCrLf & _
" DETERMINISTIC NO sql " & vbCrLf & _
"BEGIN " & vbCrLf & _
"DECLARE p1 INTEGER DEFAULT LOCATE('1h',inh); " & vbCrLf & _
"IF p1 = 0 THEN " & vbCrLf & _
" SET p1=LOCATE('1 h',inh); " & vbCrLf & _
" IF p1 = 0 THEN " & vbCrLf & _
"  SET p1=LOCATE('1 Std',inh); " & vbCrLf & _
"  IF p1 = 0 THEN " & vbCrLf & _
"   SET p1=LOCATE('1Std',inh); " & vbCrLf & _
"  END IF; " & vbCrLf & _
" END IF; " & vbCrLf & _
"END IF; " & vbCrLf & _
"IF p1 <> 0 THEN " & vbCrLf & _
" SET p1=LOCATE('mg/dl',inh,p1); " & vbCrLf & _
"ELSE " & vbCrLf & _
" SET p1=LOCATE('mg/dl',inh,locate('mg/dl',inh)+1); " & vbCrLf & _
"END IF; " & vbCrLf & _
"RETURN _vorNr(inh,p1);" & vbCrLf & _
"END"
myEFrag sql

'Glucose 2 im OGTT
sql = "DROP FUNCTION IF EXISTS `g2`"
myEFrag sql
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `g2`(inh VARCHAR(1000)) RETURNS INT(3)" & vbCrLf & _
" DETERMINISTIC NO sql " & vbCrLf & _
"BEGIN " & vbCrLf & _
"DECLARE p1 INTEGER DEFAULT LOCATE('2h',inh); " & vbCrLf & _
"IF p1 = 0 THEN " & vbCrLf & _
" SET p1=LOCATE('2 h',inh); " & vbCrLf & _
" IF p1 = 0 THEN " & vbCrLf & _
"  SET p1=LOCATE('2 Std',inh); " & vbCrLf & _
"  IF p1 = 0 THEN " & vbCrLf & _
"   SET p1=LOCATE('2Std',inh); " & vbCrLf & _
"  END IF; " & vbCrLf & _
" END IF; " & vbCrLf & _
"END IF; " & vbCrLf & _
"IF p1 <> 0 THEN " & vbCrLf & _
" SET p1=LOCATE('mg/dl',inh,p1); " & vbCrLf & _
"ELSE " & vbCrLf & _
" SET p1=LOCATE('mg/dl',inh,locate('mg/dl',inh,locate('mg/dl',inh)+1)+1);" & vbCrLf & _
"END IF; " & vbCrLf & _
"RETURN _vorNr(inh,p1);" & vbCrLf & _
"END"
myEFrag sql

sql = "DROP FUNCTION IF EXISTS `_lGFR`"
myEFrag sql
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `_lGFR`(pid INT(6)) RETURNS INT(5) " & vbCrLf & _
"    DETERMINISTIC READS SQL DATA " & vbCrLf & _
"BEGIN " & vbCrLf & _
"DECLARE muster VARCHAR(50) DEFAULT '^gfr|^gfck|^creacl'; " & vbCrLf & _
"DECLARE muc VARCHAR(50) DEFAULT '^krea02$|^krea$|^kres$|^creat$'; " & vbCrLf & _
"DECLARE eic VARCHAR(10) DEFAULT 'mg/dl'; " & vbCrLf & _
"DECLARE gfr2s,gfr1s,cre2s,cre1s CHAR(50); " & vbCrLf & _
"DECLARE gsl CHAR(1); " & vbCrLf & _
"DECLARE cre2,cre1,cre DOUBLE; " & vbCrLf & _
"DECLARE zp1, zp2, zp3, zp4 DATETIME; " & vbCrLf & _
"DECLARE p1, p2 INTEGER; " & vbCrLf & _
"-- SET @patid=pid; " & vbCrLf & _
"SELECT eingang,wert INTO zp2, gfr2s FROM laborywert w LEFT JOIN laboryus u ON u.id=w.usid WHERE pat_id = pid AND abkü RLIKE muster AND eingang=(SELECT MAX(eingang) FROM laborywert w LEFT JOIN laboryus u ON u.id=w.usid WHERE pat_id=pid AND abkü RLIKE muster) LIMIT 1; " & vbCrLf & _
"SELECT zeitpunkt,wert INTO zp1, gfr1s FROM laborneu WHERE pat_id = pid AND abkü RLIKE muster ORDER BY zeitpunkt DESC LIMIT 1; " & vbCrLf & _
"-- diese und die nächste Optimierung beschleunigen die Abfrage für die ganze aktfv von 27 auf 49s " & vbCrLf & _
"-- SELECT CAST(zeitpunkt As Date),CAST(TRIM(IF(TRIM(`n`.`Wert`) REGEXP '^[0-9]+\\,?[0-9]*$',REPLACE(`n`.`Wert`,',','.'),`n`.`Wert`)) AS char CHARSET utf8mb4) COLLATE utf8mb4_german2_ci " & vbCrLf & _
"--  INTO zp1, gfr1s FROM laborneu n" & vbCrLf & _
"--  WHERE pat_id = @patid AND CAST(`n`.`Abkü` AS char CHARSET utf8) RLIKE muster AND zeitpunkt=(SELECT MAX(CAST(zeitpunkt As Date)) FROM laborneu WHERE pat_id = @patid AND CAST(`n`.`Abkü` AS char CHARSET utf8) RLIKE muster) LIMIT 1;" & vbCrLf & _
"SELECT eingang,wert INTO zp4, cre2s FROM laborywert w LEFT JOIN laboryus u ON u.id=w.usid WHERE pat_id = pid AND abkü RLIKE muc AND einheit = eic ORDER BY eingang DESC LIMIT 1; " & vbCrLf & _
"SET cre2=zuBruch(cre2s);" & vbCrLf
sql = sql & vbCrLf & _
"-- SELECT CAST(zeitpunkt As Date),CAST(TRIM(IF(TRIM(`n`.`Wert`) REGEXP '^[0-9]+\\,?[0-9]*$',REPLACE(`n`.`Wert`,',','.'),`n`.`Wert`)) AS char CHARSET utf8mb4) COLLATE utf8mb4_german2_ci " & vbCrLf & _
"--  INTO zp3, cre1s FROM laborneu n" & vbCrLf & _
"--  WHERE pat_id = @patid AND CAST(`n`.`Abkü` AS char CHARSET utf8) RLIKE muc AND CAST(`Einheit` AS char CHARSET utf8mb4) COLLATE utf8mb4_german2_ci LIKE eic AND zeitpunkt=(SELECT MAX(CAST(zeitpunkt As Date)) FROM laborneu WHERE pat_id = @patid AND CAST(`n`.`Abkü` AS char CHARSET utf8) RLIKE muc AND CAST(`Einheit` AS char CHARSET utf8mb4) COLLATE utf8mb4_german2_ci LIKE eic) LIMIT 1;" & vbCrLf & _
"SET cre1=zuBruch(cre1s);" & vbCrLf & _
"IF zp4 IS NOT NULL AND ((zp4 > zp3 OR zp3 IS NULL) AND (zp4 > zp2 OR zp2 IS NULL) AND (zp4 > zp1 OR zp1 IS NULL)) AND cre2 IS NOT NULL AND cre2 <> 0 THEN " & vbCrLf & _
" SET cre=cre2; " & vbCrLf & _
"ELSEIF zp3 IS NOT NULL AND ((zp3 > zp2 OR zp2 IS NULL) AND (zp3 > zp1 OR zp1 IS NULL)) AND cre1 IS NOT NULL AND cre1 <> 0 THEN " & vbCrLf & _
" SET cre=cre1; " & vbCrLf & _
"ELSEIF zp2 IS NULL AND zp1 IS NULL THEN " & vbCrLf & _
" RETURN 1000; " & vbCrLf & _
"ELSEIF zp2 > zp1 OR zp1 IS NULL THEN " & vbCrLf & _
" RETURN CONVERT(zuBruch(gfr2s),INTEGER); " & vbCrLf & _
"ELSE " & vbCrLf & _
" RETURN CONVERT(zuBruch(gfr1s),INTEGER); " & vbCrLf & _
"END IF; " & vbCrLf & _
"SELECT geschlecht INTO gsl FROM namen WHERE pat_id=pid; " & vbCrLf & _
"RETURN ROUND(186*POW(cre,-1.154)*POW(patalter(pid),-0.203)*IF(gsl='w',0.742,1)); " & vbCrLf & _
"END"
myEFrag sql
' "SELECT zeitpunkt,wert INTO zp2, gfr2s FROM labor2a WHERE pat_id = pid AND abkü RLIKE muster AND ORDER BY zeitpunkt DESC LIMIT 1; " & vbCrLf & _
"RETURN 1000;" & vbCrLf & _
" SET gfr2s=TRIM(gfr2s); SET p1=INSTR(gfr2s,'/'); IF p1>0 THEN SET gfr2s=LEFT(gfr2s,p1); END IF;" & vbCrLf & _
" RETURN CONVERT(CONVERT(REPLACE(CONCAT('0',gfr2s),',','.'),DOUBLE),INTEGER); " & vbCrLf & _

sql = "DROP FUNCTION IF EXISTS `obnephrop`"
' zu koordinieren mit obLabI
' zp = Zahl pathologisch, zg = Zahl gesund, gzp = Gesamtzahl pathologisch
myEFrag sql
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `obnephrop`(pid INT(6)) RETURNS INT " & vbCrLf & _
"    DETERMINISTIC READS SQL DATA " & vbCrLf & _
"    COMMENT 'zur Beurteilung der Nephropathie' " & vbCrLf & _
"BEGIN " & vbCrLf & _
"DECLARE werth,lwerth DOUBLE;" & vbCrLf & _
"DECLARE werts,lwerts,einhh VARCHAR(100);" & vbCrLf & _
"DECLARE zph DATE;" & vbCrLf & _
"DECLARE egfr,zp,zg,gzp INTEGER(5);" & vbCrLf & _
"DECLARE obpath INTEGER(1);" & vbCrLf & _
"SELECT wert,SUM(IF(wert>=30 AND zeitpunkt>SUBDATE(NOW(),1465),1,0)),SUM(IF(wert<30 AND zeitpunkt>SUBDATE(NOW(),1465),1,0)),SUM(IF(wert>=30,1,0)),zeitpunkt,einheit INTO werts, zp,zg,gzp,zph,einhh FROM (" & vbCrLf & _
"SELECT zeitpunkt,abkü,COALESCE(wert,0) wert,einheit FROM laborneu WHERE pat_id=pid AND abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND einheit LIKE 'mg/g%' " & vbCrLf & _
"UNION " & vbCrLf & _
"SELECT eingang zeitpunkt,abkü,COALESCE(wert,0) wert,einheit FROM laborywert w LEFT JOIN laboryus u ON u.id=w.usid WHERE pat_id=pid AND abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND einheit LIKE 'mg/g%' " & vbCrLf & _
"ORDER BY wert DESC) i;" & vbCrLf & _
"SET werth=zuBruch(werts); " & vbCrLf & _
"SELECT wert INTO lwerts FROM (" & vbCrLf & _
"SELECT zeitpunkt,abkü,COALESCE(wert,0) wert,einheit FROM laborneu WHERE pat_id=pid AND abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND einheit LIKE 'mg/g%' " & vbCrLf & _
"UNION " & vbCrLf & _
"SELECT eingang zeitpunkt,abkü,COALESCE(wert,0) wert,einheit FROM laborywert w LEFT JOIN laboryus u ON u.id=w.usid WHERE pat_id=pid AND abkü IN ('ALBCRE','ALBKRE','ALBQ','ALBUM','ALBUP') AND einheit LIKE 'mg/g%' " & vbCrLf & _
"ORDER BY zeitpunkt DESC LIMIT 1) i;" & vbCrLf & _
"SET lwerth=zuBruch(lwerts);"
sql = sql & "IF (zp > 1 OR werth > 200 OR (zg=0 AND gzp>1)) AND (lwerth>=30 OR zp>2) THEN " & vbCrLf & _
"  SET obpath=1;" & vbCrLf & _
"  SET @ergeb=CONCAT(werth,' ',einhh,' (',zph,')');" & vbCrLf & _
"ELSE " & vbCrLf & _
"  SET obpath=0;" & vbCrLf & _
"  SET @ergeb='';" & vbCrLf & _
"END IF;" & vbCrLf & _
"SELECT _lGFR(pid) INTO egfr; " & vbCrLf & _
"IF egfr<=60 THEN SET obpath=1; END IF; " & vbCrLf & _
"SET @ergeb = CONCAT(@ergeb, '; ',CONVERT(egfr,CHAR(5)),' ml/min'); " & vbCrLf & _
"RETURN obpath; " & vbCrLf & _
"END"
myEFrag sql

sql = "DROP FUNCTION IF EXISTS `obbla`"
myEFrag sql
sql = "CREATE DEFINER=`praxis`@`%` FUNCTION `obbla`(inh VARCHAR(1000)) RETURNS int(1)" & vbCrLf & _
"    DETERMINISTIC " & vbCrLf & _
"    Comment 'ob Blutabnahme (aus EBM-Ziffern)' " & vbCrLf & _
"BEGIN " & vbCrLf & _
"  DECLARE zuunt VARCHAR(100); " & vbCrLf & _
"  DECLARE erg int(1) default 0; " & vbCrLf & _
"  DECLARE pos int(10); " & vbCrLf & _
"read0: LOOP " & vbCrLf & _
"  SET pos=INSTR(inh,','); " & vbCrLf & _
"  IF pos > 0 THEN " & vbCrLf & _
"   SET pos=pos-1; SET zuunt=SUBSTR(inh,1,pos); SET inh=SUBSTR(inh,pos+2); " & vbCrLf & _
"  ELSE " & vbCrLf & _
"   SET zuunt=inh;set inh=''; " & vbCrLf & _
"  END IF; " & vbCrLf & _
"  IF zuunt='' THEN LEAVE read0; END IF; " & vbCrLf & _
"  IF zuunt RLIKE '32[0-9]{3}' AND FIND_IN_SET(zuunt,'32027,32066,32237,32435,32437,32449,32720,32727,32744,')=0 THEN SET erg=1; LEAVE read0; END IF; " & vbCrLf & _
"  IF inh='' THEN LEAVE read0; END IF; " & vbCrLf & _
"END LOOP; " & vbCrLf & _
"  RETURN erg; " & vbCrLf & _
"End"
myEFrag sql



Forms(0).Ausgeb "Views für " & DBCn & " erstellt", True
syscmd 5

Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doViewsErstellen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' dbViewsErstellen

#If True Then
Sub LaborIns1(dc As Object, Pat_ID$, nurLabor%) ' nur in tuBriefStandalone
 Dim SelbstStatus%, raDatBOF%
 Dim Matr$(), MForm%(), mBreiten$(), DiffBr%, i&, j&, k&, im&
 
 On Error GoTo fehler
 Call LaborInsPLZ(Pat_ID, SelbstStatus, raDatBOF, Matr, MForm, mBreiten)
 
 Const maxbreite% = 14
 DiffBr = 0
 If UBound(Matr, 2) - 5 > maxbreite Then
   DiffBr = UBound(Matr, 2) - 5 - maxbreite
 End If
 syscmd 4, "Labor (2) nach Diffbr-Bestimmung"
 
 Dim Matr2$(), MForm2$(), obZeile%, obaltZeile%, gelZeilen&, ZusText$
 ReDim Matr2(UBound(Matr, 1), UBound(Matr, 2) - DiffBr, UBound(Matr, 3))
 ReDim MForm2(MAXvb(UBound(MForm, 2) - DiffBr, 1), UBound(Matr, 3)) ' erst: 26.7.11
 For j = 0 To UBound(Matr, 3)
'  IF j = 3 THEN Stop
  obZeile = 0
  obaltZeile = 0
  For i = UBound(Matr, 2) To 6 Step -1
   
   If Matr(0, i, j) <> vNS Then
    If i >= 6 + DiffBr Then
     obZeile = True
     Exit For
    Else
     obaltZeile = True
     Exit For
    End If
   End If
  Next i
'  obZeile = True
  If obZeile Then
   For i = 0 To UBound(Matr, 2) - DiffBr
    If i >= 6 Then im = i + DiffBr Else im = i
    If Not (gelZeilen = 0 And i = im) Or j = 0 Or True Then ' j=0 ab 4.11., da Überschrift fehlt; 22.4.12 true, da Bedingung nicht mehr verstanden
     For k = 0 To UBound(Matr, 1)
      Matr2(k, i, j - gelZeilen) = Matr(k, im, j)
      On Error Resume Next
      MForm2(i, j - gelZeilen) = MForm(im, j)
      On Error GoTo fehler
     Next k
    End If
   Next i
  Else ' => obzeile=0
   gelZeilen = gelZeilen + 1
   If obaltZeile Then
    If Matr(0, 3, j) <> vNS Then ZusText = ZusText & Matr(0, 3, j) & ", "
   End If
  End If
 Next j
 
 Select Case SelbstStatus
  Case 0
   Dim rH
   Set rH = dc.Range(dc.bookmarks("LaborÜS").Range.Start - 1, dc.bookmarks("LaborÜS").Range.Start - 1)
   dc.bookmarks("LaborÜS").Range = vNS
   dc.bookmarks.Add name:="LaborÜS", Range:=dc.Range(rH.Start + 1, rH.Start + 1)
  Case 1
   dc.bookmarks("LaborÜS").Range = "Laborbefunde (nur mitgeteilte, woanders bestimmte Werte):"
  Case 2
   dc.bookmarks("LaborÜS").Range = "Bei uns erhobene Laborbefunde:"
 End Select
 syscmd 4, "Labor (2) nach Zeilenzahlbestimmung"
 
 
 If raDatBOF Then
    dc.bookmarks!LaborÜS.Range = vNS
    dc.bookmarks!labhinw.Range = vNS
    Exit Sub
 End If
 syscmd 4, "Labor (2) nach Überschrift"
 
' dc.Application.Visible = True
 If nurLabor = 0 Then Wapp.ScreenUpdating = False
' ZZ+1 ; rz+3
If UBound(Matr, 3) - gelZeilen + 1 > 0 Then
 Set Tabl = dc.Tables.Add(Range:=dc.bookmarks!laborTab.Range, NumRows:=UBound(Matr, 3) - gelZeilen + 1, NumColumns:=UBound(Matr, 2) - 2 - DiffBr, _
               DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitContent)
 dc.bookmarks.Add name:="LaborZusatz", Range:=dc.Range(Tabl.Range.END, Tabl.Range.END)
' dc.Range(Tabl.Range.End, Tabl.Range.End) = "Vor dem " & Matr2(0, 6, 0) & " wurden noch bestimmt: " & ZusText
 If LenB(ZusText) > 2 Then
  ZusText = Left$(ZusText, Len(ZusText) - 2)
  dc.bookmarks!LaborZusatz.Range = "Vor dem " & Matr2(0, 6, 0) & " wurden noch bestimmt: " & ZusText & Chr(13)
  dc.Range(dc.bookmarks("LaborZusatz").Range.paragraphs(1).Range.Start, dc.bookmarks("LaborZusatz").Range.paragraphs(1).Range.END).Font.size = 9
  With dc.Range(Tabl.Range.END, Tabl.Range.END).ParagraphFormat
   .LeftIndent = CentimetersToPoints(IIf(UBound(Matr, 2) - 5 > 14, -1.9, 0))
   .FirstLineIndent = CentimetersToPoints(0)
  End With
 End If
 syscmd 4, "Labor (2) nach Tabellen-Addition"
 
 dc.bookmarks.Add name:="DMP", Range:=dc.Range(Tabl.Range.END, dc.bookmarks!DMP.Range.END)
 syscmd 4, "Labor (2) nach DMP-Markenlöschung"
 With Tabl
  If WappBuild > 9 Then
   If .style <> "Tabellengitternetz" And .style <> "Tabellenraster" Then ' Word 19 = Tabellenraster
     .style = "Tabellengitternetz"
   End If
  .ApplyStyleHeadingRows = True
  .ApplyStyleLastRow = True
  .ApplyStyleFirstColumn = True
  .ApplyStyleLastColumn = True
  End If
  .Rows(1).Range.Orientation = 1
  .Rows(1).Range.Font.bold = True
  syscmd 4, "Labor (2) nach TabellenStil"
  For j = 1 To UBound(Matr, 3) - gelZeilen + 1 'ZZ ' -1 Gestrichen am 24.1.12 ' +1 seit 10.12.12
   'If .Cell(j, 1) <> vns THEN
   .cell(j, 1).Range.Font.bold = True
  Next j
'   .Rows.LeftIndent = CentimetersToPoints(-0.7)
' Seitenrand links
   .Rows.LeftIndent = CentimetersToPoints(IIf(UBound(Matr, 2) - 5 > 14, -2, -1))
   .PreferredWidth = CentimetersToPoints(19.4)
 End With
 syscmd 4, "Labor (2) nach Tabellenformatierung"
 
'  For j = 0 To UBound(Matr, 3)
'   For i = 6 To UBound(Matr, 2)
'    IF (MForm(i, j) AND 2) = 2 THEN
'     MForm(3, j) = 2 ' Parameterspalte mit pathologischem Wert fett auf gelbem Hintergrund
'     Exit For
'    END IF
'   Next i
'   AusS.AppVar (Array(" <tr>", vbCrLf))
'   Dim iumw& ' um Laborwerte in umgekehrter Datumsreihenfolge anzuzeigen
'   For i = 3 To UBound(Matr, 2)
'    IF i > 5 THEN
'     iumw = UBound(Matr, 2) - i + 6
'    Else
'     iumw = i
'    END IF
'    Dim Formg$ ' Formatierung / Farbe
'    SELECT CASE MForm(iumw, j)
'     Case 0: Formg = "lnn" ' Labor normal eigen
'     Case 1: Formg = "lno" ' Labor normal importiert
'     Case 2: Formg = "lpn" ' Labor pathologisch eigen
'     Case 3: Formg = "lpo" ' Labor pathologisch importiert
'    END SELECT
'   Next i
'  Next j
 
 For i = 3 To UBound(Matr2, 2) ' rz+5
  For j = 0 To UBound(Matr2, 3) - gelZeilen  ' ZZ ' -1 gestrichen am 24.1.12
    If i <= UBound(MForm2, 1) Then
     If MForm2(i, j) <> vNS Then
      If (MForm2(i, j) And 1) = 1 Then
       Tabl.cell(j + 1, i - 2).Range.Font.Italic = True
      End If
'    IF (MForm(i, j) AND 2) = 2 THEN
      If MForm2(i, j) > 1 Then
       Tabl.cell(j + 1, i - 2).Range.Font.bold = True  'V3
      End If
     End If
    End If
   If Len(Matr2(0, i, j)) < 11 Or i < 6 Then
    Tabl.cell(j + 1, i - 2) = Matr2(0, i, j)
   Else
    Tabl.cell(j + 1, i - 2) = Left$(Matr2(0, i, j), 8) & ".."
   End If
  Next j
 Next

' For i = 3 To UBound(Matr, 2) - DiffBr ' rz+5 - DiffBr
'  For j = 0 To UBound(Matr, 3) - 1 ' ZZ
'   Dim im&
'   im = i
'   IF i >= 6 THEN
'    im = im + DiffBr
'   END IF
'    IF (MForm(im, j) AND 1) = 1 THEN
'     Tabl.cell(j + 1, i - 2).Range.Font.Italic = True
'    END IF
'    IF (MForm(im, j) AND 2) = 2 THEN
'     Tabl.cell(j + 1, i - 2).Range.Font.Bold = True
'    END IF
'   IF Len(Matr(0, im, j)) < 11 OR i + DiffBr < 6 THEN
'    Tabl.cell(j + 1, i - 2) = Matr(0, im, j)
'   Else
'    Tabl.cell(j + 1, i - 2) = Left$(Matr(0, im, j), 8) & ".."
'   END IF
'  Next j
' Next
 syscmd 4, "Labor (2) nach Dateneintragung "
 
' For j = 2 To Tabl.Rows.Count
'  For i = 1 To Tabl.Columns.Count
'   IF i <> 2 AND LenB(Trim$(replace$(replace$(Tabl.cell(j, i).Range.Text, vbCr, vNS), Chr$(7), vNS))) <> 0 THEN GoTo Nr
'  Next i
'  Tabl.Rows(j).Delete
'  j = j - 1
' Next j
'Nr:
 
 syscmd 4, "Labor (2) vor Endformatierung"

 With Tabl
  Select Case UBound(Matr, 2) - 5
   Case 1, 2, 3, 4, 5, 6
     .Range.Font.size = 9
'     .Columns(1).Width =
'     .Columns(2).Width = 53
'     .Columns(3).Width = 44
   Case 7, 8, 9, 10, 11
     .Range.Font.size = 8
'     .Columns(1).Width = 96.6
'     .Columns(2).Width = 49
'     .Columns(3).Width = 41
   Case 12, 13, 14
     .Range.Font.size = 7
'     .Columns(1).Width = 85.6
'     .Columns(2).Width = 43.5
'     .Columns(3).Width = 36.5
   Case Else
     .Range.Font.size = 6
'     .Columns(1).Width = 23
'     .Columns(2).Width = 48
'     .Columns(3).Width = 43
  End Select
 End With
' Weiten:
' 6: 73   / 38   / 33
' 7: 85,6 / 43,5 /36,5
' 8: 96,6 / 49 / 41
' 9:      / 53 / 44
 syscmd 4, "Labor (2) nach Endformatierung"
 End If
#If False Then
  Dim adn$, dcn$, Tn$
  adn = Wapp.activedocument.name
  dcn = dc.name
  Tabl.id = "Labor"
  Set Wapp = Nothing
  GetWord
  Set dc = Wapp.documents(dcn)
  Dim tabli
  For Each tabli In dc.Tables
   If tabli.id = "Labor" Then
    Set Tabl = tabli
    Exit For
   End If
  Next tabli
 #End If
syscmd 4, "Labor (2) vor LaborTabAnp"
'Do While True
Call LaborTabAnp(Wapp, dc, nurLabor)
'Loop
syscmd 4, "Labor (2) nach LaborTabAnp"
 
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborIns1/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' LaborIns1

#Else
'Sub LaborIns1(dc AS Object, Pat_id&, nurLabor%) 'Word.Document, Pat_id&)
' Dim raLw As New ADODB.Recordset, raDat As New ADODB.Recordset, rLaU As New ADODB.Recordset, ls$
'' Zeilenzahl bestimmen
' Dim ZZ&, rz$, gschl$, Vgl$, altGruppe%, Nb$ ' Normbereich
' Dim sql1$
' Dim u!, o! ' oberer und unterer Grenzwert numerisch
' Dim uNG$, oNG$ ' obere und untere Normgrenze in Zeichen
' Dim unm$, onm$, unw$, onw$
' Dim pKz$ ' Pathologisch-Kennzeichen
' Dim SelbstStatus
' Dim Matr$(), mBreiten$()
' Dim altgru$
' ON Error GoTo fehler
'' sql = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" AS NB FROM (" & laborAbfr & " WHERE pat_id = " & Pat_id & ") AS labor WHERE not (Wert LIKE ""%eine Berechnung%"") " & _
'  "UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar,Normbereich AS NB " + _
'  "FROM `laborxus` LEFT JOIN laborxwert ON laborxus.RefNr=laborxwert.RefNr " + _
'  "WHERE pat_id = " + CStr(Pat_id) + " AND NOT (Wert LIKE ""%eine Berechnung%"") AND NOT EXISTS (SELECT * FROM `laborneu` WHERE pat_id = " + CStr(Pat_id) + " AND abkü = laborxwert.Abkü AND wert = laborxwert.wert AND zeitpunkt > laborxus.Eingang-3 AND zeitpunkt < laborxus.Eingang+6)"
'' sql = "SELECT * FROM labor1 WHERE pat_id = " & Pat_id & " AND NOT wert LIKE '%eine Berechnung%' ORDER BY zeitpunkt DESC UNION SELECT * FROM labor2 WHERE pat_id = " & Pat_id & " AND NOT wert LIKE '%eine Berechnung%' ORDER BY zeitpunkt DESC"
'' sql = "SELECT * FROM labor1 WHERE pat_id = " & Pat_id & " AND NOT wert LIKE '%eine Berechnung%' UNION SELECT * FROM labor2 WHERE pat_id = " & Pat_id & " AND NOT wert LIKE '%eine Berechnung%'"
'' sql = "SELECT * FROM `labor2a` WHERE pat_id = " & Pat_id & " AND NOT wert LIKE '%eine Berechnung%' GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_id & " AND NOT wert LIKE '%eine Berechnung%' "
' sql = "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_id & " AND NOT wert LIKE '%eine Berechnung%' UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_id & " AND NOT wert LIKE '%eine Berechnung%') i GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb"
'
' ON Error GoTo fehler
' SelbstStatus = 0
'' SET rDat = Dtb.OpenRecordset("SELECT * FROM `" + QMdbAkt + "`.labor WHERE pat_id = " & CStr(Pat_id) & " AND kommentar = ""Dieser Eintrag wurde manuell erzeugt.""")
' sql1 = "SELECT * FROM (" & laborAbfr & " WHERE pat_id = " & Pat_id & ") AS labor WHERE kommentar = ""Dieser Eintrag wurde manuell erzeugt."""
' raDat.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
' IF Not raDat.BOF THEN SelbstStatus = SelbstStatus + 1
'' SET rDat = Dtb.OpenRecordset("SELECT * FROM `" + QMdbAkt + "`.labor WHERE pat_id = " & CStr(Pat_id) & " AND kommentar <> ""Dieser Eintrag wurde manuell erzeugt.""")
' sql1 = "SELECT * FROM (" & laborAbfr & " WHERE pat_id = " & Pat_id & ") AS labor WHERE kommentar <> ""Dieser Eintrag wurde manuell erzeugt."""
' SET raDat = Nothing
' raDat.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
' IF Not raDat.BOF THEN SelbstStatus = SelbstStatus + 2
' SELECT CASE SelbstStatus
'  Case 0
'   Dim rH
'   SET rH = dc.Range(dc.Bookmarks("LaborÜS").Range.Start - 1, dc.Bookmarks("LaborÜS").Range.Start - 1)
'   dc.Bookmarks("LaborÜS").Range = vNS
'   dc.Bookmarks.Add Name:="LaborÜS", Range:=dc.Range(rH.Start + 1, rH.Start + 1)
'  Case 1
'   dc.Bookmarks("LaborÜS").Range = "Laborbefunde (nur mitgeteilte, woanders bestimmte Werte):"
'  Case 2
'   dc.Bookmarks("LaborÜS").Range = "Bei mir erhobene Laborbefunde:"
' END SELECT
'
' syscmd 4, "Labor nach Überschrift"
'
' gschl = vNS
' altGruppe = 0
'' SET rLW = TabÖff("Namen", "Pat_id")
'' ralw.Seek "=", Pat_id
' myFrag raLw, "SELECT geschlecht FROM `namen` WHERE pat_id = " & Pat_id
' IF Not raLw.EOF THEN
'  gschl = raLw!Geschlecht
' END IF
' raLw.Close
' Dim sql0$
' Dim sql2$
'nochmal:
' sql2 = "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_id & " UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_id & ") i GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb"
' sql0 = "SELECT DISTINCT gruppe, reihe, sql1.abkü, sql1.einheit FROM (" + sql2 + ") AS sql1 LEFT JOIN `laborparameter` ON sql1.abkü = `laborparameter`.abkü AND iIF(ISNULL(sql1.einheit) OR sql1.einheit="""", ""kA"",sql1.einheit) = `laborparameter`.einheit WHERE pat_id = " + CStr(Pat_id) + " AND (NOT ISNULL(wert) OR NOT ISNULL(kommentar)) AND (reihe <> 999 OR ISNULL(reihe)) ORDER BY gruppe, reihe"
' IF lies.obMySQL THEN sql0 = replace$(sql0, "iIF(", "IF(")
' SET raLw = Nothing
'' GoTo nochmal
' raLw.Open sql0, DBCn, adOpenDynamic, adLockReadOnly
' altgru = vNS
' Do While Not raLw.EOF
'  IF raLw!gruppe <> altgru THEN
'   ZZ = ZZ + 1
'   altgru = raLw!gruppe
'  END IF
'  ZZ = ZZ + 1
'  raLw.MoveNext
' Loop
'
' syscmd 4, "Labor nach Zeilenzahlbestimmung"
'' sql1 = "SELECT COUNT(0) FROM (SELECT DISTINCT abkü FROM (" + sql + ") AS sql1 WHERE (NOT ISNULL(wert) OR NOT ISNULL(kommentar))) AS innen"
''' sql = "SELECT COUNT(0) FROM (SELECT DISTINCT abkü FROM `" + QMdbAkt + "`.laborUNION WHERE pat_id = " + CStr(Pat_id) + " AND (NOT ISNULL(wert) OR NOT ISNULL(kommentar)))"
''' SET raLW = Dtb.OpenRecordset(sql1)
'' SET raLW = Nothing
'' raLW.Open sql1, dbcn, adOpenDynamic, adLockReadOnly
'' zZ = raLW.Fields(0)
'' raLW.Close
'' sql1 = "SELECT COUNT(0) FROM (SELECT DISTINCT gruppe FROM (" + sql + ") AS sql1 LEFT JOIN `laborparameter` ON sql1.abkü = `laborparameter`.abkü AND iIF(ISNULL(sql1.einheit) OR sql1.einheit="""", ""kA"",sql1.einheit) = `laborparameter`.einheit WHERE pat_id = " + CStr(Pat_id) + " AND (NOT ISNULL(wert) OR NOT ISNULL(kommentar))) AS innen"
''' sql = "SELECT COUNT(0) FROM (SELECT DISTINCT gruppe FROM (SELECT * FROM `" + QMdbAkt + "`.laborUNION LEFT JOIN `" + QMdbAkt + "`.`laborparameter` ON laborunion.abkü = `laborparameter`.abkü AND iIF(ISNULL(laborunion.einheit) OR laborunion.einheit="""", ""kA"",laborunion.einheit) = `laborparameter`.einheit WHERE pat_id = " + CStr(Pat_id) + " AND (NOT ISNULL(wert) OR NOT ISNULL(kommentar))))"
''' SET rLW = Dtb.OpenRecordset(sql1)
'' IF lies.obmysql THEN sql1 = replace$(sql1, "iIF(", "IF(")
'' SET raLW = Nothing
'' raLW.Open sql1, dbcn, adOpenDynamic, adLockReadOnly
'' zZ = zZ + raLW.Fields(0)
'' raLW.Close
'nochmal1:
' sql1 = "SELECT DISTINCT * FROM (SELECT DISTINCT DATE(zeitpunkt) AS Datum FROM (" + sql + ") AS sql1 LEFT JOIN `laborparameter` ON sql1.abkü = `laborparameter`.abkü AND iIF(ISNULL(sql1.einheit) OR sql1.einheit="""", ""kA"",sql1.einheit) = `laborparameter`.einheit WHERE pat_id = " + CStr(Pat_id) + " AND (NOT ISNULL(wert) OR NOT ISNULL(kommentar)) AND (reihe <> 999 OR ISNULL(reihe))) AS i0 ORDER BY datum"
'' das erste DISTINCT ist für Access nötig
'' SET rDat = Dtb.OpenRecordset("SELECT COUNT(0) FROM (" + sql1 + ")")
'' rZ = radat.Fields(0) ' dateserial(YEAR(zeitpunkt),month(zeitpunkt),day(zeitpunkt))
' SET raDat = Nothing
' IF lies.obMySQL THEN
'  sql1 = replace$(replace$(sql1, "datevalue(", "date("), "iIF(", "IF(")
' END IF
'' GoTo nochmal1
' raDat.Open "SELECT COUNT(0) AS ct FROM (" + sql1 + ") AS innen", DBCn, adOpenDynamic, adLockReadOnly
' IF Not raDat.BOF THEN
'  rz = raDat!ct
' END IF
' Const MaxBreite% = 17
' Dim DiffBr%
' DiffBr = 0
' IF rz > MaxBreite THEN
'   DiffBr = rz - MaxBreite
' END IF
' raDat.Close
' ReDim Matr(1, rz + 5, ZZ + 1) ' links noch: 0:Abkü, 1:unterer NW, 2:oberer NW, dann 3:Langtext, 4:Einheit, 5:Normbereich
' ReDim mBreiten(rz)
'
' syscmd 4, "Labor nach Diffbr-Bestimmung"
'
' SET raDat = Nothing
' raDat.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
'' SET rDat = Dtb.OpenRecordset(sql1)
' IF raDat.BOF THEN
'    dc.Bookmarks!LaborÜS.Range = vNS
'    dc.Bookmarks!LabHinw.Range = vNS
'    Exit Sub
' END IF
' Dim i&, j&
' i = 1
'' raDat.Move DiffBr
' Do While Not raDat.EOF
'  Matr(0, i + 5, 0) = Format$(raDat!Datum, "dd.mm.yy")
'  i = i + 1
'  raDat.Move 1
'  IF i > rz THEN Exit Do
' Loop
'
' syscmd 4, "Labor nach Datumseintragung"
'
' Dim rlp As New ADODB.Recordset
' j = 1
' altgru = vNS
' raLw.MoveFirst
' Do While Not raLw.EOF
'  IF raLw!gruppe <> altgru THEN
'   altgru = raLw!gruppe
'   Matr(0, 0, j) = "Gruppe " & raLw!gruppe & "|" & raLw!reihe
''   Debug.Print Matr(0, j) & "," & Matr(1, j) & "," & Matr(2, j)
'   j = j + 1
'  END IF
'  Matr(0, 0, j) = raLw!Abkü
'  Matr(0, 4, j) = raLw!Einheit
'  SET rlp = Nothing
'  Dim rLpgeht%
'  rLpgeht = -1
'  myFrag rlp, "SELECT * FROM `laborparameter` WHERE abkü = '" & raLw!Abkü & "' AND einheit = '" & raLw!Einheit & "' AND (unm <> """" OR unw <> """" OR onm <> """" OR onw <> """")"
'  IF rlp.BOF THEN rLpgeht = 0
'  IF rLpgeht THEN IF ISNULL(rlp!Langtext) THEN rLpgeht = 0
'  IF rLpgeht THEN IF rlp!Langtext = vNS THEN rLpgeht = 0
'  IF rLpgeht THEN
'    Matr(0, 3, j) = rlp!Langtext
'  Else
'   Dim rLangt As New ADODB.Recordset
'   SET rLangt = Nothing
'   myFrag rLangt, "SELECT LangText FROM (" & laborAbfr & " WHERE abkü = '" & raLw!Abkü & "' AND einheit = '" & raLw!Einheit & "' AND NOT ISNULL(langtext) AND langtext <> """") AS labor"
'   IF rLangt.BOF THEN
'    SET rLangt = Nothing
'    myFrag rLangt, "SELECT LangText FROM (" & laborAbfr & " WHERE abkü = '" & raLw!Abkü & "' AND NOT ISNULL(langtext) AND langtext <> """") AS labor"
'   END IF
'   IF Not rLangt.BOF THEN
'    Matr(0, 3, j) = rLangt!Langtext
'   END IF
'  END IF
'  IF Not rlp.EOF THEN
'   IF gschl = "m" OR (rlp!unw = vNS AND rlp!onw = vNS) THEN
'    Matr(0, 5, j) = rlp!unm & "-" & rlp!onm
'    Matr(0, 1, j) = replace$(IIF(ISNULL(rlp!unm), vNS, rlp!unm), "1:", vNS)
'    Matr(0, 2, j) = replace$(IIF(ISNULL(rlp!onm), vNS, rlp!onm), "1:", vNS)
'   Else
'    Matr(0, 5, j) = rlp!unw & "-" & rlp!onw
'    Matr(0, 1, j) = replace$(IIF(ISNULL(rlp!unw), vNS, rlp!unw), "1:", vNS)
'    Matr(0, 2, j) = replace$(IIF(ISNULL(rlp!onw), vNS, rlp!onw), "1:", vNS)
'   END IF
'   IF Matr(0, 0, j) LIKE "LDL*" THEN Matr(0, 5, j) = "-100"
'  END IF
''  Debug.Print Matr(0, j) & "," & Matr(1, j) & "," & Matr(2, j)
'  j = j + 1
'  raLw.Move 1
' Loop
'
' syscmd 4, "Labor nach Normbereichseingabe"
'
'' dc.Application.Visible = True
' IF nurLabor = 0 THEN Wapp.ScreenUpdating = False
' SET Tabl = dc.Tables.Add(Range:=dc.Bookmarks!laborTab.Range, NumRows:=ZZ + 1, NumColumns:=rz + 3 - DiffBr, _
'               DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:=wdAutoFitContent)
' syscmd 4, "Labor nach Tabellen-Addition"
' dc.Bookmarks.Add Name:="DMP", Range:=dc.Range(Tabl.Range.End, dc.Bookmarks!DMP.Range.End)
' syscmd 4, "Labor nach DMP-Markenlöschung"
' With Tabl
'  IF WappBuild > 9 THEN
'   IF .Style <> "Tabellengitternetz" THEN
'     .Style = "Tabellengitternetz"
'   END IF
'  .ApplyStyleHeadingRows = True
'  .ApplyStyleLastRow = True
'  .ApplyStyleFirstColumn = True
'  .ApplyStyleLastColumn = True
'  END IF
'  .Rows(1).Range.Orientation = 1
'  .Rows(1).Range.Font.Bold = True
'  syscmd 4, "Labor nach TabellenStil"
'  For j = 1 To ZZ
'   'If .Cell(j, 1) <> vns THEN
'   .cell(j, 1).Range.Font.Bold = True
'  Next j
'   .Rows.LeftIndent = CentimetersToPoints(-0.7)
' END With
'
' syscmd 4, "Labor nach Tabellenformatierung"
'
' SET raLw = Nothing
' raLw.Open sql, DBCn, adOpenDynamic, adLockReadOnly
' Do While Not raLw.EOF
'  i = 6
'  Do While CDate(Matr(0, i, 0)) < DateValue(raLw!Zeitpunkt)
'   i = i + 1
'   IF i >= rz + 5 THEN Exit Do
'  Loop
'  j = 0
'  IF i > UBound(Matr, 2) THEN
'   i = UBound(Matr, 2)
'   MsgBox "Stop in LaborIns: " & vbCrLf & "i: " & i & vbCrLf & "ubound(matr,2): " & UBound(Matr, 2)
'   Stop
'  END IF
'  Do While j <= UBound(Matr, 3)
'   IF Matr(0, 0, j) = raLw!Abkü THEN
'    IF Matr(0, 4, j) = raLw!Einheit THEN
'     Exit Do
'    END IF
'   END IF
'   j = j + 1
'  Loop
'  IF j > UBound(Matr, 3) THEN j = UBound(Matr, 3)
'  Matr(0, i, j) = raLw!Wert
'  IF LenB(raLw!Kommentar) <> 0 THEN Matr(1, i, j) = raLw!Kommentar
'  IF NOT ISNULL(raLw!Kommentar) AND raLw!Kommentar <> vNS THEN
'   IF InStrB(raLw!Kommentar, "manuell") > 0 THEN
'     ON Error Resume Next
'     Tabl.cell(j + 1, i - 2).Range.Font.Italic = True
'     ON Error GoTo fehler
'   ElseIf raLw!Wert = vNS THEN
''     Matr(i + Diffbr, j) = Matr(i, j) & IIF(Matr(i, j) = vns, vns, " ") & raLW!Kommentar
'     Matr(0, i, j) = raLw!Kommentar
'   END IF
'  END IF
'  pKz = "n" ' normal
'  IF IsNumeric(raLw!Wert) OR raLw!Wert LIKE "1:*" OR raLw!Wert LIKE "<*" OR raLw!Wert LIKE ">*" THEN
'   Dim rlWS$, rlWZ#, obkl%, obgr%
'   Dim buch$, pos&
'   obkl = 0
'   obgr = 0
'   rlWS = raLw!Wert
'   For pos = 1 To Len(raLw!Wert)
'    IF InStrB("0123456789,.:-+<>", Mid$(rlWS, pos, 1)) = 0 THEN
'     rlWS = LEFT(rlWS, pos - 1) & Mid$(rlWS, pos + 1)
'     pos = pos - 1
'    END IF
'   Next pos
'   IF rlWS LIKE "1:<*" THEN rlWS = "<1:" & Mid$(rlWS, 4)
'   IF LEFT(rlWS, 1) = "<" THEN
'    obkl = True
'    rlWS = Mid$(rlWS, 2)
'   ElseIf LEFT(rlWS, 1) = ">" THEN
'    obgr = True
'    rlWS = Mid$(rlWS, 2)
'   END IF
'   rlWS = replace$(replace$(replace$(replace$(rlWS, "^", vNS), "1:", vNS), ".", ","), ",,", ",")
'   IF rlWS = vNS THEN rlWS = "0"
'   rlWZ = CDbl(rlWS)
'   IF LenB(Matr(0, 1, j)) <> 0 THEN
'    IF rlWZ < CDbl(replace$(Matr(0, 1, j), ".", ",")) AND NOT obgr THEN
'     pKz = "p"
'    END IF
'   END IF
'   IF LenB(Matr(0, 2, j)) <> 0 THEN
'    ON Error Resume Next
'    IF rlWZ > CDbl(replace$(Matr(0, 2, j), ".", ",")) AND NOT obkl THEN
'     pKz = "p"
'    END IF
'    ON Error GoTo fehler
'   END IF
''   IF o <> -99 AND rlwd > o THEN pKz = "p"
'  END IF
'  ON Error Resume Next
'  IF pKz = "p" THEN 'If Tabl.Cell(j + 1, i - 2) <> vns THEN
'   Tabl.cell(j + 1, i - 2).Range.Font.Bold = True
'  END IF
'  ON Error GoTo fehler
''  IF ISNULL(raLW!gruppe) THEN
''   altgruppe = 0
''  Else
''   altgruppe = Val(raLW!gruppe)
''  END IF
'  raLw.MoveNext
' Loop
'' dc.Bookmarks!laborTab.Range
' 'sql1 = "SELECT * FROM (SELECT distinct(cdate(int(zeitpunkt))) AS Datum FROM (" + sql + ") AS sql LEFT JOIN `" + QMdbAkt + "`.`laborparameter` ON sql.abkü = `laborparameter`.abkü AND iIF(ISNULL(sql.einheit) OR sql.einheit="""", ""kA"",sql.einheit) = `laborparameter`.einheit WHERE pat_id = " + CStr(Pat_id) + " AND (NOT ISNULL(wert) OR NOT ISNULL(kommentar))) ORDER BY datum"
'' SET rLW = Dtb.OpenRecordset(sql1, dbOpenDynaset)
'
'syscmd 4, "Labor nach Dateneintragung "
'
'' Exit Sub
' For i = 3 To rz + 5 - DiffBr
'  For j = 0 To ZZ
'   IF i = 3 THEN
'    IF Matr(0, 3, j) = "HbA1c Eigenlabor" THEN
'     Matr(0, 3, j) = "HbA1c"
'    END IF
'   END IF
'   Tabl.cell(j + 1, i - 2) = IIF(Len(Matr(0, i + IIF(i < 6, 0, DiffBr), j)) < 11 OR i + DiffBr < 6, Matr(0, i + IIF(i < 6, 0, DiffBr), j), LEFT(Matr(0, i + IIF(i < 6, 0, DiffBr), j), 8) & "..")
'  Next j
' Next
'
' For j = 2 To Tabl.Rows.Count
'  For i = 1 To Tabl.Columns.Count
'   IF i <> 2 AND Trim$(replace$(replace$(Tabl.cell(j, i).Range.Text, vbCr, vNS), Chr$(7), vNS)) <> vNS THEN GoTo Nr
'  Next i
'  Tabl.Rows(j).Delete
'  j = j - 1
' Next j
'Nr:
' ' ls = CStr(rZ) & vbcrlf
'' raDat.MoveFirst
'' Dim CNr%
'' CNr = rZ + 3
'' raDat.MoveLast
'' Do While Not raDat.BOF
''  ' ls = LS + format$(radat!Datum, "dd.mm.yyyy") & vbcrlf
''  IF lies.obmysql THEN
''   Tabl.cell(1, CNr).Range = raDat!Datum 'BDTtoDate(mid$(raDat!Datum, 7, 2) & mid$(raDat!Datum, 5, 2) & LEFT(raDat!Datum, 4))
''  Else
''   Tabl.cell(1, CNr).Range = format$(raDat!Datum, "dd.mm.yy")
''  END IF
''  CNr = CNr - 1
''  raDat.Move -1
'' Loop
' Tabl.cell(1, 1).Range = "Parameter"
' Tabl.cell(1, 2).Range = "Einheit"
' Tabl.cell(1, 3).Range = "Normbereich"
'' Exit Sub
'' ' ls = LS + CStr(zZ) & vbcrlf
'' sql1 = "SELECT * FROM (SELECT reihe, gruppe, unw, onw, unm, onm, sql1.*, DATE(zeitpunkt) AS Datum, sql1.abkü AS sabkü, sql1.einheit AS seinheit, sql1.langtext AS slangtext FROM (" + sql + ") AS sql1 LEFT JOIN `laborparameter` ON sql1.abkü = `laborparameter`.abkü AND iIF(ISNULL(sql1.einheit) OR sql1.einheit="""", ""kA"",sql1.einheit) = iIF(ISNULL(`laborparameter`.einheit),"""",`laborparameter`.einheit) WHERE pat_id = " + CStr(Pat_id) + " AND (NOT ISNULL(wert) OR NOT ISNULL(kommentar))) AS i0 ORDER BY gruppe,reihe,datum" ' dateserial(YEAR(zeitpunkt),month(zeitpunkt),day(zeitpunkt))
'' IF lies.obmysql THEN sql1 = replace$(replace$(sql1, "datevalue(", "date("), "iIF(", "IF(")
'' SET raLW = Nothing
''' SET rLW = Dtb.OpenRecordset(sql1)
'' raLW.Open sql1, dbcn, adOpenDynamic, adLockReadOnly
'' Dim aktZ%, VorWert$, AktGru$, AktSp%, AktAbkü$, obgeändert%
'' Do While Not raLW.EOF
''  IF NOT ISNULL(raLW!sabkü) THEN
''   IF ISNULL(raLW!gruppe) THEN
'''    IF rLaU Is Nothing THEN SET rLaU = Dtb.OpenRecordset("SELECT * FROM (" + sql + ") AS sql ORDER BY zeitpunkt")
''    SET raLau = Nothing
''    raLau.Open "SELECT *,sql.abkü AS sabkü, sql.langtext AS slangtext FROM (" & sql & ") AS sql ORDER BY zeitpunkt", dbcn, adOpenDynamic, adLockReadOnly
''    Debug.Print "Kein Parametereintrag für: " + CStr(raLW!sabkü) & " " & CStr(IIF(ISNULL(raLW!slangtext), vns, raLW!slangtext)) + " `" + IIF(ISNULL(raLW!seinheit), vns, raLW!seinheit) + "`"
''    Call raLau.Find("Pat_id = " & CStr(raLW!Pat_id) & " AND zeitpunkt = cdate(""" & CStr(raLW!Zeitpunkt) & """) AND fertigstgrad = """ & CStr(raLW!FertigStGrad) & """ AND abkü = """ & CStr(raLW!sabkü) & """", 0, adSearchForward, 1)
'
''    IF Not rLaU.BOF THEN
''     Call LaborParameter(raLau)
''     obgeändert = True
''    END IF
''   END IF
''  END IF
''  raLW.Move 1
'' Loop
'' IF obgeändert THEN
''  SET raLW = Nothing
'''  SET raLW = Dtb.OpenRecordset(sql)
''  raLW.Open sql, dbcn, adOpenDynamic, adLockReadOnly
'' Else
'' raLW.MoveFirst
'' END IF
'' aktZ = 1
''  Dim TestFeld$
''  Dim TestInh
''  ON Error Resume Next
''  Err.Clear
''  TestFeld = "sql."
''  TestInh = raLW.Fields(TestFeld + "abkü")
''  IF Err.Number <> 0 THEN
''   TestFeld = vns
''   Err.Clear
''  END IF
''  ON Error GoTo fehler
'' Do While Not raLW.EOF
''  IF NOT ISNULL(raLW.Fields(TestFeld + "abkü")) THEN
'''   IF Val(ralw!Gruppe) <> altGruppe THEN
''    ' ls = LS & vbcrlf
'''   END IF
''   IF ISNULL(VorWert) OR VorWert = "" OR raLW.Fields(TestFeld + "LangText") <> VorWert THEN
''    aktZ = aktZ + 1
''    AktSp = 4
''    ON Error GoTo gruppenfehler
''    IF NOT ISNULL(AktGru) AND AktGru <> "" AND raLW.Fields("Gruppe") <> AktGru THEN
''    ON Error GoTo fehler
''      aktZ = aktZ + 1
''    END IF
''    Do While aktZ > Tabl.Rows.Count
''     Tabl.Rows.Add
''     zZ = zZ + 1
''    Loop
''    AktGru = IIF(ISNULL(raLW.Fields("Gruppe")), vns, raLW.Fields("Gruppe"))
''    Tabl.cell(aktZ, 1).Range = raLW.Fields("slangtext") 'IIF(ISNULL(raLW.Fields("sql.LangText")), vns, raLW.Fields("sql.LangText"))
''    Tabl.cell(aktZ, 2).Range = raLW.Fields("seinheit") 'IIF(ISNULL(raLW.Fields("sql.Einheit")), vns, raLW.Fields("sql.Einheit"))
''    u = -99
''    o = -99
''    uNG = vns
''    oNG = vns
''    unw = IIF(ISNULL(raLW!unw), vns, raLW!unw)
''    onw = IIF(ISNULL(raLW!onw), vns, raLW!onw)
''    unm = IIF(ISNULL(raLW!unm), vns, raLW!unm)
''    onm = IIF(ISNULL(raLW!onm), vns, raLW!onm)
'''    IF instrb(raLW.Fields("sql.Langtext"), "LDL") > 0 THEN
''    IF instrb(raLW.Fields("slangtext"), "LDL") > 0 THEN
''     onm = "100"
''    END IF
''    IF gschl = "w" THEN
''     uNG = replace$(unw, "1:", "")
''     oNG = replace$(onw, "1:", "")
''    END IF
''    IF uNG = "" AND oNG = "" THEN
''     uNG = replace$(unm, "1:", "")
''     oNG = replace$(onm, "1:", "")
''    END IF
''    IF uNG <> "" THEN u = CDbl(uNG)
''    IF oNG <> "" THEN o = CDbl(oNG)
'''   IF IsNumeric(raLw!Wert) OR raLw!Wert LIKE "1:*" THEN
'''    IF gschl = "w" THEN
'''     IF Not ralw!unw = "" THEN u = CDbl(replace$(ralw!unw, "1:", ""))
'''     IF Not ralw!onw = "" THEN o = CDbl(replace$(ralw!onw, "1:", ""))
'''    END IF
'''    IF gschl = "m" OR u = -99 THEN
'''     IF Not ralw!unm = "" THEN u = CDbl(replace$(ralw!unm, "1:", ""))
'''    END IF
'''    IF gschl = "m" OR o = -99 THEN
'''     IF Not ralw!onm = "" THEN o = CDbl(replace$(ralw!onm, "1:", ""))
'''    END IF
'''   END IF
''    ' ls = LS + pKz & vbcrlf
''    NB = vns
''    IF uNG <> "" OR oNG <> "" THEN
''     IF gschl = "w" AND (unw <> "" OR onw <> "") THEN
''      NB = unw + "-" + onw
''     END IF
''    END IF
''    IF NB = "" AND (unm <> "" OR onm <> "") THEN
''     NB = unm + "-" + onm
''    END IF
''    Tabl.cell(aktZ, 3) = NB
''    ' ls = LS + NB & vbcrlf
''   END IF
'''   VorWert = IIF(ISNULL(raLW.Fields("sql.LangText")), vns, raLW.Fields("sql.LangText"))
''   VorWert = raLW.Fields("slangtext") 'IIF(ISNULL(raLW.Fields("sql.LangText")), vns, raLW.Fields("sql.LangText"))
''   ' ls = LS + ralw.Fields("sql.LangText") & vbcrlf
''   ' ls = LS + ralw.Fields("sql.Einheit") & vbcrlf
''   ' ls = LS + format$(ralw!Datum, "dd.mm.yyyy") & vbcrlf
''   ' ls = LS + ralw!Wert & vbcrlf
'
''   Do
''    IF DateValue(raLW!Datum) > CDate(LEFT(Tabl.cell(1, AktSp).Range, 8)) THEN
''     AktSp = AktSp + 1
''    Else
''     Exit Do
''    END IF
''    IF AktSp > rZ + 3 THEN
''     Exit Do
''    END IF
''   Loop
'
''   Tabl.cell(aktZ, AktSp).Range = IIF(ISNULL(raLW!Wert), IIF(ISNULL(raLW!Kommentar), vns, format$(raLW!Zeitpunkt, "dd/mm/yy") + ": " + raLW!Kommentar), replace$(replace$(trim$(IIF(ISNULL(raLW!Wert), vns, raLW!Wert)), vbcr, ""), vbtab, ""))
''   IF instrb(raLW!Kommentar, "manuell") > 0 THEN Tabl.cell(aktZ, AktSp).Range.Font.Italic = True
'
''  END IF ' NOT ISNULL(ralw!`sql!abkü`
''  raLW.Move 1
'' Loop '  While Not ralw.EOF
'
'' Open "c:\wwtest.txt" For Output AS #34
'' Print #34, LS
'' Close #34
' syscmd 4, "Labor vor Endformatierung"
'
' With Tabl
'  SELECT CASE rz
'   Case 1, 2, 3, 4, 5, 6
'     .Range.Font.size = 9
''     .Columns(1).Width =
'     .Columns(2).Width = 53
'     .Columns(3).Width = 44
'   Case 7, 8, 9, 10, 11
'     .Range.Font.size = 8
'     .Columns(1).Width = 96.6
'     .Columns(2).Width = 49
'     .Columns(3).Width = 41
'   Case 12, 13, 14
'     .Range.Font.size = 7
'     .Columns(1).Width = 85.6
'     .Columns(2).Width = 43.5
'     .Columns(3).Width = 36.5
'   Case Else
'     .Range.Font.size = 6
'     .Columns(1).Width = 73
'     .Columns(2).Width = 38
'     .Columns(3).Width = 33
'  END SELECT
'' Weiten:
'' 6: 73   / 38   / 33
'' 7: 85,6 / 43,5 /36,5
'' 8: 96,6 / 49 / 41
'' 9:      / 53 / 44
' syscmd 4, "Labor nach Endformatierung"
' raDat.Close
' raLw.Close
' END With
'
' #If False THEN
'  Dim adn$, dcn$, Tn$
'  adn = Wapp.activedocument.Name
'  dcn = dc.Name
'  Tabl.ID = "Labor"
'  SET Wapp = Nothing
'  GetWord
'  SET dc = Wapp.documents(dcn)
'  Dim tabli
'  For Each tabli In dc.Tables
'   IF tabli.ID = "Labor" THEN
'    SET Tabl = tabli
'    Exit For
'   END IF
'  Next tabli
' #END IF
'
'syscmd 4, "Labor vor LaborTabAnp"
' Call LaborTabAnp(Wapp, dc)
'syscmd 4, "Labor nach LaborTabAnp"
'
' Exit Sub
'gruppenfehler:
' IF Err.Number = 3265 THEN ' Gruppe fehlt in Auflistung, 18.4.06
'  MsgBox "Resume in LaborIns1: Err.Nummer = 3265"
'  Resume
'  Resume Next
' END IF
' Exit Sub
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIF(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborIns1/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End Sub      ' LaborIns1
#End If

Sub LaborTabAnp(Wapp As Object, dc As Object, nurLabor%) 'Word.Application, dc AS Word.Document)
 Dim i%, j%, k%, Tz%, wi#, wineu!, mk%, buch$
 Dim FontSize%
 On Error GoTo fehler
 If dc.Tables.COUNT > 0 Then
 If nurLabor = 0 Then Wapp.ScreenUpdating = False
 With dc.Tables(dc.Tables.COUNT)
 FontSize = .cell(1, 1).Range.Font.size
  .AllowAutoFit = False
  For i = 1 To .Columns.COUNT
   wi = 0
   mk = 0
   For j = 2 To .Rows.COUNT
'    wineu = (.Cell(j, i).Range.Characters.Count)
'    For k = wineu To 0 Step -1
'     tz = Asc(mid$(.Cell(j, i).Range, k, 1))
'     IF tz < 32 THEN
'      wineu = wineu - 1
'     Else
'      Exit For
'     END IF
'    Next
'   wineu = wineu+1 ' das Viereck (chr$(7)) am Schluß
' schneller, wenngleich nicht ganz so robust:
    wineu = .cell(j, i).Range.Characters.COUNT
    wineu = FontSize * wineu
    If wineu > wi Then
     If i > 3 Then
      If InStrB(.cell(j, i).Range, ",") = 0 And InStrB(.cell(j, i).Range, ".") = 0 Then
        mk = 0
      Else
        mk = -1
      End If
     End If
     wi = wineu
 '    Debug.Print "Längenrekord Spalte ", i, ":", .Cell(j, i).Range.Characters.Count, "(" + .Cell(j, i).Range + ")"
    ElseIf wineu = wi And mk = -1 Then
      If InStrB(.cell(j, i).Range, ",") = 0 And InStrB(.cell(j, i).Range, ".") = 0 Then
       mk = 0
      End If
    End If
   Next j
'   Debug.Print "Spalte: ", i, wi
   Select Case i
    Case 1
     wi = wi * 0.41  ' 0.5 * (3.41 / 3.21) ' paßt
    Case 2
     wi = wi * 0.61   ' 0.5 * (1.74 / 1.36) ' paßt
    Case 3
     wi = wi * 0.62   ' 0.5 * (1.47 / 1.23) ' paßt
    Case Is > 3
      If mk = -1 Then
       wi = 0# + ((wi - 1) * 0.78)
      Else
       wi = wi * 0.88  '1 * 0.5 + (wi - 1) * 1
      End If
   End Select
   If i > 3 Then
    If wi > 40 Then wi = 40
   End If
   Select Case i
'    Case 1: wi = 82.5
'    Case 2: wi = 32.55
'    Case 3: wi = 44.2
'    Case Else: wi = 30
   End Select
'   .Columns.AutoFit
   If wi < 7.2 Then wi = 7.2
   .Columns(i).Width = wi
'   For j = 1 To .Rows.Count
'    .Cell(j, i).Range.FitTextWidth = 0
'   Next j
  Next
'  .Rows.LeftIndent = CentimetersToPoints(0)
'  .Columns.AutoFit
'  .Rows.LeftIndent = CentimetersToPoints(0)

' .PreferredWidthType = 3 'wdPreferredWidthPoints
 Dim breite!, spnr%
 For spnr = 1 To .Columns.COUNT
  breite = breite + .Columns(spnr).Width
 Next spnr
 Const maxbreite! = 19
 If breite > CentimetersToPoints(maxbreite) Then .PreferredWidth = CentimetersToPoints(maxbreite)
End With
 Wapp.ScreenUpdating = True
 End If
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborTabAnp/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' LaborTabAnp

Sub LaborTabAnp_alt(dc)
 Dim diff%, erl%, aend%, i%, D1, D2
 Dim ColuNr%
 On Error GoTo fehler
 With Tabl
  For ColuNr = 1 To .Columns.COUNT
'   Debug.Print "vorher:  " + CStr(ColuNr) & " " & CStr(.Columns(ColuNr).Width)
  Next ColuNr
'  d1 = Now + 2 / 60 / 60 / 24
'  Do While d2 < d1
'   d2 = Now
'  Loop
' Das folgende funktioniert leider nicht im Gegensatz zum Obigen
  Tabl.Select
  Do While .Columns(.Columns.COUNT - 1).Width = .Columns(.Columns.COUNT).Width And .Columns(.Columns.COUNT - 1).Width = .Columns(1).Width
  Loop
'  For ColuNr = 1 To .Columns.COUNT
'   Debug.Print "vorher2: " + CStr(ColuNr) & " " & CStr(.Columns(ColuNr).Width)
'  Next ColuNr
  For ColuNr = 4 To .Columns.COUNT
   If .Columns(ColuNr).Width > 80 Then GoTo ender
  Next ColuNr
'  For ColuNr = 1 To .Columns.COUNT
'   Debug.Print "nachher: " + CStr(ColuNr) & " " & CStr(.Columns(ColuNr).Width)
'  Next ColuNr
  GoTo nichtender
ender:
  For i = 2 To 3
   erl = .Columns(ColuNr).Width - 80
   aend = IIf(i = 2, 58, 56) - .Columns(i).Width
   If erl < aend Then aend = erl
   .Columns(i).Width = .Columns(i).Width + aend
   .Columns(ColuNr).Width = .Columns(ColuNr).Width - aend
  Next
 End With
nichtender:
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborTabAnp_alt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' LaborTabAnp
#If False Then
Sub Tabelle88888888() ' Scheint wohl nicht vorzukommen, da Verunstaltung möglich
Dim i%, zZStr$, ZZ%, rZStr$, rz%, diff%, ct$ ' Container
Dim aktZ% ' aktuelle Zeile
Dim par$, Einh$, Datu$, Wert$, path$, Nb$, altPar$
Dim F As File
With Wapp.activedocument
 If .Tables.COUNT > 0 Then
  With .Tables(.Tables.COUNT)
    If FSO Is Nothing Then Set FSO = CreateObject("Scripting.FileSystemObject")
    Set F = FSO.GetFile(aVerz + "\Labor.txt") ' uverz & "Anamnese"
    If Not F Is Nothing Then
      Close #16
      Open F.path For Input As #16
      Line Input #16, rZStr
      rz = Val(rZStr)
      diff = rz - .Columns.COUNT + 3
      If diff > 0 Then
        For i = 1 To diff
         .Columns.Add
        Next
      ElseIf diff < 0 Then
       For i = 1 To -diff
        .Columns(.Columns.COUNT).Delete
       Next
      End If
      For i = 1 To rz ' Datum
       Line Input #16, ct
       .cell(1, 3 + i).Range = Left$(ct, 10) ' mit ganzem Jahr
      Next
      Line Input #16, zZStr
      ZZ = Val(zZStr)
      diff = ZZ - .Rows.COUNT
      If diff > 0 Then
       For i = 1 To diff
        .Rows.Add
       Next
      ElseIf diff < 0 Then
       For i = 1 To -diff
        .Rows(.Rows.COUNT).Delete
       Next
      End If
      With .Range
        .Font.bold = False
        .Font.Italic = False
        .Font.Color = RGB(0, 0, 0) ' wdfontblack
      End With
      altPar = vNS
      Line Input #16, ct ' obligatorische Leerzeile
      aktZ = 1
      Do While Not EOF(16)
       Line Input #16, par
       If par = "" Then ' neue Gruppe
        aktZ = aktZ + 1
        For i = 1 To .Columns.COUNT
         .cell(aktZ, i).Range = vNS
        Next i
        Line Input #16, par
       End If
       If par <> altPar Then
        aktZ = aktZ + 1
        For i = 4 To .Columns.COUNT
         .cell(aktZ, i).Range = vNS
        Next i
        altPar = par
       End If
       .cell(aktZ, 1).Range = par
       Line Input #16, Einh
       .cell(aktZ, 2).Range = Einh
       Line Input #16, Datu
       For i = 4 To .Columns.COUNT
        If Left$(.cell(1, i).Range, 10) = Datu Then
         Exit For
        End If
       Next i
       Line Input #16, Wert
       .cell(aktZ, i).Range = Wert
       Line Input #16, path
       If path = "p" Then
        .cell(aktZ, i).Range.Font.bold = True
        .cell(aktZ, i).Range.Font.Italic = True
        .cell(aktZ, i).Range.Font.Color = wdColorRed
       Else
       End If
       Line Input #16, Nb
       .cell(aktZ, 3).Range = Nb
      Loop
      Close #16
    End If ' not f is nothing
    
'    Exit Sub
  
        .Columns(1).Width = 92
        .Columns(2).Width = 47
        .Columns(3).Width = 42
        For i = 4 To .Columns.COUNT
         .cell(1, i).Range = Format$(Left$(.cell(1, i).Range, 10), "dd.mm.yyyy")
         .Columns(i).Width = 30
        Next i
        For i = 2 To .Rows.COUNT
         If .cell(i, 3).Range.paragraphs.COUNT > 2 Then
          .cell(i, 3).Range = vNS
         End If
        Next i
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderHorizontal)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderVertical)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    With .Range.Find
        .clearformatting
        .replacement.clearformatting
        .Text = "Eiwei¯"
        .replacement.Text = "Eiweiß"
        .Execute REPLACE:=wdReplaceAll
        .clearformatting
        .replacement.clearformatting
        .Text = "Á"
        .replacement.Text = "µ"
        .Execute REPLACE:=wdReplaceAll
        .clearformatting
        .Font.Color = wdColorRed
        .Text = vNS
        With .replacement
          .clearformatting
          .Font.bold = True
          .Font.Italic = True
          .Text = vNS
        End With
       .Execute Format:=True, REPLACE:=wdReplaceAll
    End With ' .range.find
  .Rows(1).Range.Font.bold = True
  For i = 1 To .Rows.COUNT
   .cell(i, 1).Range.Font.bold = True
  Next i
'  With .Columns(1)
'   .Range.Font.Bold = True
''   .SELECT
''   SELECTion.Font.Bold = True
''   ActiveDocument.Range(Start:=.Cells(1).Range.Start, End:=.Cells(.Cells.Count).Range.End).Font.Bold = True
 ' END With
'  IF 1 = 0 THEN
'   Dim wrkJet AS Workspace, Adr AS Database, rLab AS DAO.Recordset, rNam AS DAO.Recordset
'   Dim fehler ' wg. Kompilierfehler bei untenstehender Zeile trotz Einbzug aller Verweise wie in winword
   ' Public fehler As New C_Fehlerbehandlung
'   IF fehler.Ist_TMAktiv THEN
'    SET wrkJet = CreateWorkspace$("", "Admin", vns, dbUseJet)
'    SET Adr = wrkJet.OpenDatabase(uverz & "Anamneseblatt.mdb", False, False)
'    SET rNam = Adr.OpenRecordset("Namen", dbOpenTable)
'    rNam.index = "Auswahl"
'    SET TMInterface.TMTxt.Daten.AktivesDokument = ActiveDocument
'    rNam.Seek "=", TMInterface.TMTxt.Daten.AktuellerPatient.Namensdaten.Nachname, TMInterface.TMTxt.Daten.AktuellerPatient.Namensdaten.Vorname, TMInterface.TMTxt.Daten.AktuellerPatient.Geburtsdaten.Datum
'    IF Not rNam.NoMatch THEN
'     SET rLab = Adr.OpenRecordset("SELECT * FROM `" + QMdbAkt + "`.Labor WHERE pat_id = " + CStr(rNam!Pat_id) + ";", dbOpenDynaset)
'     Do While Not rLab.EOF
''     IF Not IsNumeric(rLab!Wert) AND NOT ISNULL(rLab!Wert) AND rLab!Wert <> vns THEN
'       Debug.Print rLab!LangText & " " & CStr(rLab!Wert) & " " & rLab!Einheit
''     END IF
'      rLab.Move 1
'     Loop
'    END IF
'   END IF
'  END IF
  End With ' .tables
  End If
 End With
 With Wapp.options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
 End With

 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Tabelle/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' Tabelle
#End If
'Function Verlauf$(Pat_id&)
' Verlauf = eintraege(Pat_id, """notiz"",""telef"",""med""")
'End FUNCTION ' Verlauf$(Pat_id&)
Function eintraege$(Pat_ID$, krit$, Optional VorDat As Date)
#Const obAlte = True
 Dim raVL As New ADODB.Recordset, lzp As Date, aktdat As Date
 On Error GoTo fehler
' Call dtbInit
 Select Case LCase$(krit)
  Case """rr"""
'   SET raVL = Dtb.OpenRecordset("SELECT zeitpunkt,rr AS inhalt FROM `" + QMdbAkt + "`.rr WHERE pat_id = " & pat_id + " ORDER BY zeitpunkt")
'   raVL.Open "SELECT zeitpunkt,""rr"" art, IF(ISNULL(bemerkung) OR bemerkung='',rr,CONCAT(rr,' (',bemerkung,')')) inhalt FROM rr WHERE pat_id = " & Pat_id + " ORDER BY zeitpunkt", DBCn, adOpenDynamic, adLockReadOnly
   myFrag raVL, "SELECT zeitpunkt,""rr"" art, IF(ISNULL(bemerkung) OR bemerkung='',rr,CONCAT(rr,' (',bemerkung,')')) inhalt FROM rr WHERE pat_id = " & Pat_ID + " ORDER BY zeitpunkt"
  Case """rr"",""puls"""
'   raVL.Open "SELECT * FROM (SELECT zeitpunkt, 'rr' art, IF(ISNULL(bemerkung) OR bemerkung='',rr, CONCAT(rr,' (',bemerkung,')')) inhalt FROM rr WHERE pat_id = " & Pat_id & " UNION SELECT zeitpunkt,art,inhalt FROM `eintraege` WHERE pat_id = " & Pat_id & " AND art IN (""rr"",""puls"")) innen ORDER BY zeitpunkt", DBCn, adOpenDynamic, adLockReadOnly
   myFrag raVL, "SELECT * FROM (SELECT zeitpunkt, 'rr' art, IF(ISNULL(bemerkung) OR bemerkung='',rr, CONCAT(rr,' (',bemerkung,')')) inhalt FROM rr WHERE pat_id = " & Pat_ID & " UNION SELECT zeitpunkt,art,inhalt FROM `eintraege` WHERE pat_id = " & Pat_ID & " AND art IN (""rr"",""puls"")) innen ORDER BY zeitpunkt"
Case Else
'   SET raVL = Dtb.OpenRecordset("SELECT zeitpunkt,inhalt,art FROM `" + QMdbAkt + "`.`eintraege` WHERE pat_id = " & pat_id + " AND art IN (" & Krit & ") ORDER BY zeitpunkt")
#If Not obAlte Then
'   raVL.Open "SELECT zeitpunkt,inhalt,art FROM `eintraege` WHERE pat_id = " & Pat_id + " AND art IN (" & krit & ") AND zeitpunkt > " & DatFor_k(VorDat) & " ORDER BY zeitpunkt", DBCn, adOpenDynamic, adLockReadOnly
   myFrag raVL, "SELECT zeitpunkt,inhalt,art FROM `eintraege` WHERE pat_id = " & Pat_ID + " AND art IN (" & krit & ") AND zeitpunkt > " & DatFor_k(VorDat) & " ORDER BY zeitpunkt"
#Else
   myFrag raVL, "SELECT zeitpunkt,inhalt,art FROM `eintraege` WHERE pat_id = " & Pat_ID + " AND art IN (" & krit & ") ORDER BY zeitpunkt"
#End If
 End Select
 Do While Not raVL.EOF
  Dim art$, Inhalt$, Zp$, Übs$
  art = raVL!art
  Inhalt = raVL!Inhalt
  Zp = raVL!Zeitpunkt
  If Inhalt <> "Blutentnahme" And Inhalt <> "Blutabnahme" Then
   aktdat = DateValue(Zp)
   If aktdat <> lzp Then
    eintraege = eintraege + Format$(aktdat, "DD.MM.YY")
   End If
   If LCase$(krit) <> """rr""" And art = "usdm" Then
    eintraege = eintraege + vbTab + REPLACE$(REPLACE$(Inhalt, "^", vNS) & vbCrLf, "aktuellen Blutdruck und ggf. Puls bitte extra eingeben", vNS)
   Else
    Übs = ""
    Select Case raVL!art
     Case "ulcus": Übs = "Ulcus: "
     Case "wv": Übs = "Wundverband: "
     Case "debr": Übs = "Debridement: "
     Case "fuß": Übs = "Fußbefund: "
     Case "kv": Übs = "Kompressionsverband: "
     Case "fa", "fam": Übs = "Familienanamnese: "
     Case "rauch": Übs = "Rauchanamnese: "
     Case "bew", "beweg", "bewegung", "bewg": Übs = "Bewegung: "
     Case "alko": Übs = "Alkoholanamnese: "
     Case "colo": Übs = "Colo-/Gastroskopien: "
     Case "aug", "augen": Übs = "Augen: "
     Case "beruf": Übs = "Beruf: "
    End Select
    Inhalt = Übs & Inhalt
    eintraege = eintraege + vbTab + REPLACE$(Inhalt, "^", vNS) & vbCrLf
   End If
   lzp = DateValue(Zp)
  End If
  raVL.Move 1
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in eintraege/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' eintraege$(Pat_id&, Krit$)

'Function haakt() ' Hausärzte aus dem Internet ergänzen (Knopf)
' Dim rHa AS DAO.Recordset, geht%
' ON Error GoTo fehler
' SET rHa = TabÖff("Hausaerzte", "KVNr")
' Do While Not rHa.EOF
''  IF rHA!Nachname = "Licht" OR geht THEN
'   Call do_haakt(rHa)
''   geht = -1
''  END IF
'  rHa.Move 1
' Loop
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIF(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in haakt/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' haakt()
#If mitab Then
Function do_ha_click(ByVal HANr)
#If False Then
 Dim stelle%
 Dim rHa As DAO.Recordset
 If KVÄDatei = vNS Then Call KVÄDateifind
 Set kvä = DBEngine.OpenDatabase(KVÄDatei, , True)
 Set rHa = kvä.OpenRecordset("HAe", dbOpenTable)
 rHa.Index = "KVNr"
' Unterprogramm
 Dim teile$()
 teile = Split(HANr, "_")
 If UBound(teile) > -1 Then
  If InStrB(teile(0), "/") = 0 Then
   teile(0) = Left$(teile(0), 2) + "/" + Mid$(teile(0), 3) 'replace$(teile(0), "/", vns)
  End If
 End If
 If UBound(teile) = 2 Then rHa.Seek "=", teile(0), Trim$(teile(1)), Trim$(teile(2))
 If UBound(teile) = 1 Or rHa.Nomatch Then rHa.Seek "=", teile(0), Trim$(teile(1))
 If UBound(teile) = 0 Or rHa.Nomatch Then rHa.Seek "=", teile(0)
 Dim stAppName$, DBNr$
 If Not rHa.Nomatch Then DBNr = rHa!DBNr
 If DBNr <> vNS And Not IsNull(DBNr) Then
  stAppName = Environ("programfiles") + "\internet explorer\iexplore.exe http://www.kvb.de/servlet/PB/cmd/arztsuche3/index.html?&enr=" + CStr(DBNr)
  Call Shell(stAppName, vbNormalFocus)
 End If
 kvä.Close
#End If
End Function
#End If
'Function KVÄDateifind$()
' Dim Fil AS File
' Dim FGef AS File
' Dim db AS DAO.Database
' Dim rs AS DAO.Recordset, Sta AS DAO.Recordset
' Dim altStD#
' ON Error GoTo fehler
' IF FSO Is Nothing THEN SET FSO = New FileSystemObject
' For Each Fil In FSO.GetFolder(aVerz).Files
'  IF Fil.Name LIKE "KV*.mdb" THEN
'   ON Error Resume Next
'   SET db = DAO.OpenDatabase(Fil.Path)
'   SET rs = db.OpenRecordset("HAE")
'   rs.MoveLast
'   rs.MoveFirst
'   IF rs.RecordCount > 14000 THEN
'    SET Sta = db.OpenRecordset("SELECT * FROM Stand ORDER BY stand DESC")
'    Sta.Move 30
'    IF altStD = 0 OR (altStD > 0 AND NOT Sta.EOF AND NOT Sta.BOF AND Sta!datum > altStD) THEN
'     altStD = Sta!datum
'     SET FGef = Fil
'    END IF
'   END IF
'   ON Error GoTo fehler
'  END IF
' Next Fil
' KVÄDatei = FGef.Path
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = currentDB.Name
'#Else
' AnwPfad = App.Path
'#END IF
'SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIF(ISNULL(Err.source), vns, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in KVÄDateifind/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): Progende
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End SELECT
'End FUNCTION ' KVÄDateiFind

' in Halokal
Function do_haakt(rHa As ADODB.Recordset)
 Dim ars As New ADODB.Recordset, aRsx As New ADODB.Recordset
' Call KVÄVorb
' Call acon(HaT)
 If LenB(DBCn) = 0 Or DBCn = "" Then Call acon(quelleT)
 Dim sql$
 sql = "SELECT * FROM `kvaerzte`.`hae` WHERE kvnr = '" & REPLACE$(rHa!KVNr, "/", vNS) & "' AND nachname = '" & rHa!Nachname & "' AND vorname = '" & rHa!Vorname & "'"
 myFrag ars, sql 'haecn
 If ars.BOF Then
  ars.Close
  sql = "SELECT * FROM `kvaerzte`.`hae` WHERE kvnr = '" & REPLACE$(rHa!KVNr, "/", vNS) & "' AND nachname = '" & rHa!Nachname & "'"
  myFrag ars, sql 'haecn
  If ars.BOF Then
   sql = "SELECT * FROM `kvaerzte`.`hae` WHERE nachname = '" & rHa!Nachname & "' AND vorname = '" & rHa!Vorname & "' AND ort = '" & rHa!ort & "'"
   myFrag ars, sql 'haecn
  End If
 End If
 If ars.BOF Then
  MsgBox "Stop durch ars.Bof Vermutlich veraltete KV-Nummer in Turbomed: " & rHa!KVNr & " " & rHa!Nachname & " " & rHa!Vorname & vbCrLf & "sql: " & sql
  Stop
 Else
  MsgBox "Stop durch not ars.Bof: " & rHa!KVNr & " " & rHa!Nachname & " " & rHa!Vorname & vbCrLf & "sql: " & sql
  Stop
'     rHa.Edit
     rHa!name = IIf(IsNull(ars!anrede), vNS, ars!anrede + " ") + IIf(IsNull(ars!Titel), vNS, ars!Titel + " ") + IIf(IsNull(ars!Vorname), vNS, ars!Vorname + " ") + IIf(IsNull(ars!Nachname), vNS, ars!Nachname)
     rHa!anschrift = ars!Straße & ", " & ars!plz & " " & ars!ort
     rHa!telefon = ars!tel1
     If Not IsNull(ars!fax1) And ars!fax1 <> vNS Then
      rHa!telefax = ars!fax1
     Else
      myFrag aRsx, "SELECT * FROM `kvaerzte`.`haxls` WHERE name = '" & ars!Nachname & "' AND vorname = '" & ars!Vorname & "' AND kvnr = '" & REPLACE$(ars!KVNr, "/", vNS) & "'" 'haecn
      If aRsx.BOF Then
       myFrag aRsx, "SELECT * FROM `kvaerzte`.`haxls` WHERE name = '" & ars!Nachname & "' AND vorname = '" & ars!Vorname & "'" 'haecn
      End If
      If aRsx.BOF Then
       MsgBox "Vermutlich SchreibFehler in Turbomedtabelle bei: " & ars!KVNr & " " & ars!Nachname & " " & ars!Vorname
      Else
       If Not IsNull(aRsx!fax) And ars!fax <> vNS Then
        rHa!telefax = ars!fax
       End If
      End If
     End If
     If Not (IsNull(ars!email) Or ars!email = vNS) Then rHa!e_mail = ars!email
     rHa!zulassungsgebiet = ars!zulg
     rHa!arzttyp = ars!arzttyp
     rHa!dmpt2 = ars!dmpt2
     rHa!beme = ars!beme
     If IsNull(rHa!Vorname) Or rHa!Vorname = vNS Then rHa!Vorname = ars!Vorname
     If IsNull(rHa!Nachname) Or rHa!Nachname = vNS Then rHa!Nachname = ars!Nachname
     rHa!Titel = ars!Titel
     rHa!geschlecht = ars!geschlecht
     rHa!ort = ars!ort
     rHa!plz = ars!plz
     rHa!Straße = ars!Straße
     rHa!KVNr = Left$(REPLACE$(ars!KVNr, "/", vNS), 7)
     rHa.Update
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_haakt/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_haakt(rHa AS dao.recordset)

Function PiepKurz()
 On Error GoTo fehler
 Call WD
 Call Sound(WinDir + "\media\Windows XP-Hinweis.wav")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PiepKurz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' PiepKurz()
Function Piep()
 On Error GoTo fehler
 Call WD
 Call Sound(WinDir + "\media\ringout.wav")
 Sleep 1000
 Call Sound(WinDir + "\media\recycle.wav")
 Sleep 1000
 Call Sound(WinDir + "\media\ringout.wav")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Piep/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Piep()

#If False Then
' 19.6.11: wird offenbar nicht verwendet
Function GetZA(DBNr&)
         Dim hClipMemory&
         Dim lpClipMemory&
         Dim MyString$
         Dim RetVal&
 On Error GoTo fehler
         If OpenClipboard(0&) = 0 Then
            Call PiepKurz
            Call MsgBox("Konnte Zwischenspeicher nicht öffnen. Könnte besetzt sein", , Titel)
            Exit Function
         End If

         ' Obtain the handle to the global memory
         ' block that is referencing the text.
         hClipMemory = GetClipboardData(CF_TEXT)
         If IsNull(hClipMemory) Then
            Call PiepKurz
            Call MsgBox("Konnte keinen Speicher reservieren", , Titel)
            GoTo OutOfHere
         End If

         ' Lock Clipboard memory so we can reference
         ' the actual data string.
         lpClipMemory = GlobalLock(hClipMemory)

         If Not IsNull(lpClipMemory) Then
            MyString = Space$(MAXSIZE)
            RetVal = lstrcpy(MyString, lpClipMemory)
            RetVal = GlobalUnlock(hClipMemory)

            ' Peel off the null terminating character.
            If InStrB(MyString, vbNullChar) > 0 Then
             MyString = Mid$(MyString, 1, InStr(1, MyString, vbNullChar, 0) - 1)
            End If
         Else
            Call PiepKurz
            Call MsgBox("Konnte Speicher nicht reservieren, um den String zu kopieren", , Titel)
         End If

OutOfHere:

         RetVal = CloseClipboard()
         GetZA = MyString
         
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Call PiepKurz
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + " bei DS Nr. " + CStr(IIf(IsNull(DBNr), vNS, DBNr)) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetZA/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GetZA

'schneidet Neue-Zeile-Zeichen ab
Function z$(rHa, FeName$, ByVal q, DBNr&)
 On Error GoTo fehler
 Select Case rHa.Fields(FeName).Type
  Case 10
   z = Left$(q, rHa.Fields(FeName).size)
  Case Else ' 12 = Memo
   z = q
 End Select
 
  GoTo schluss:
schluss:
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Call PiepKurz
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + " bei DS Nr. " + CStr(DBNr) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Z(neueZeileZeichen abschneiden)/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Z$
#End If

Function snie()
  Call WD
  Call Sound(WinDir + "\media\Windows XP-Batterie niedrig.wav")
End Function ' snie()

Function WD() ' WinDir
 On Error GoTo fehler
 WinDir = Space$(144)
 Call GetWindowsDirectory(WinDir, 144)
 WinDir = REPLACE$(Trim$(WinDir), vbNullChar, vNS)
 WD = WinDir
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in WD/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' WD

#If False Then
Function testdb()
 Dim anam As DAO.Recordset
 On Error GoTo fehler
 Set testdb = OpenDatabase(QMdbAkt)
 Set anam = testdb.OpenRecordset("Anamnesebogen", dbOpenTable)
 anam.Index = "Pat_id"
 anam.Seek "=", 97
 anam.Edit
 anam.Update
 testdb.Close
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in testdb/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' testdb
#End If
Function Datenbankkontrolle()
 Dim obGleich%, pText$
' Dim haz AS DAO.Recordset, hae AS DAO.Recordset, hae1 AS DAO.Recordset
 Dim haz As ADODB.Recordset, hae As ADODB.Recordset, hae1 As ADODB.Recordset
 Dim kvärs As ADODB.Recordset
' Dim kvä AS DAO.Database
 On Error GoTo fehler
' Call dtbInit
 ' IF kvä is nothing THEN
' IF Not lies.obmysql AND KVÄDatei = vns THEN Call KVAccSuch
' Call KVÄVorb
'  Call acon(HaT)
 If LenB(DBCn) = 0 Or DBCn = "" Then Call acon(quelleT)
' SET kvä = DBEngine.OpenDatabase(KVÄDatei, , True)
' kvä.Open "kvaerzte.hae", DBCn, adOpenDynamic, adLockReadOnly 'haecn
' myFrag kvärs, "SELECT * FROM kvaerzte.hae" ' auskommentiert 1.6.24
 Open aVerz + "\HAFehler.txt" For Output As #32 ' uverz & "Anamnese"
' SET haz = TabÖff("Hausaerzte")
' haz.Open "hausaerzte", DBCn, adOpenDynamic, adLockReadOnly
 myFrag haz, "SELECT * FROM hausaerzte"
' SET hae = kvä.OpenRecordset("Hae", dbOpenTable)
' hae.Open "kvaerzte.hae", DBCn, adOpenDynamic, adLockReadOnly 'haecn
 myFrag hae, "SELECT * FROM kvaerzte.hae"
' SET hae1 = kvä.OpenRecordset("Hae", dbOpenDynaset)
' hae1.Open "kvaerzte.hae", DBCn, adOpenDynamic, adLockReadOnly 'haecn
 myFrag hae1, "SELECT * FROM kvaerzte.hae"
' hae.index = "KVNr"
 Do While Not haz.EOF
'  hae.Seek "=", LEFT(haz!KVNr, 2) + "/" + right$(haz!KVNr, 5) ' "kvnr = " +
  Set hae = Nothing
'  hae.Open "SELECT * FROM `kvaerzte`.`hae` WHERE kvnr = " & haz!kvnu
  myFrag hae, "SELECT * FROM `kvaerzte`.`hae` WHERE kvnr = " & haz!kvnu
  If hae.BOF Then
   pText = CStr(haz!id) + ": " + CStr(IIf(IsNull(haz!KVNr), "KV-Nr: (Null)", haz!KVNr)) & " " & IIf(IsNull(haz!Nachname), "Nachname: (Null)", haz!Nachname) & " " & IIf(IsNull(haz!Vorname), "Vorname: (Null)", haz!Vorname) + " in HAE nicht gefunden"
'   hae1.FindFirst "instr(HAName, """ + haz!Nachname + """) > 0 AND instr(HAName, """ + haz!Vorname + """) > 0 AND ort = """ + haz!Ort + """"
   Set hae1 = Nothing
   myFrag hae1, "SELECT * FROM `kvaerzte`.`hae` WHERE haname LIKE '%" & haz!Nachname & "%' AND haname LIKE '%" & haz!Vorname & "%' AND ort LIKE '%" & haz!ort & "%'" 'haecn
   If Not hae1.EOF Then
    pText = pText + ",        dort aber " + haz!Nachname & ", " & haz!Vorname + " in " + haz!ort + " mit KV-Nr: " + hae1!KVNr + "!"
   End If
   Print #32, pText
  Else
   obGleich = 0
   Do
    If hae!Nachname = haz!Nachname And hae!Vorname = haz!Vorname Then
     obGleich = -1
     Exit Do
    End If
    hae.Move 1
    If hae.EOF Then Exit Do
    If hae!KVNr <> REPLACE$(haz!KVNr, "/", vNS) Then Exit Do
   Loop
   If Not obGleich Then
    Print #32, CStr(haz!id) + ": " + CStr(IIf(IsNull(haz!KVNr), "KV-Nr: (Null)", haz!KVNr)) & " " & IIf(IsNull(haz!Nachname), "Nachname: (Null)", haz!Nachname) & " " & IIf(IsNull(haz!Vorname), "Vorname: (Null)", haz!Vorname) + " in HAE nicht mit gleichem Namen nicht gefunden"
   End If
  End If
  haz.Move 1
 Loop
 Close #32
' kvärs.Close ' auskommentiert 1.6.24
' Dtb.Close
 MsgBox "Fertig mit Datenbankkontrolle!"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Datenbankkontrolle/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Datenbankkontrolle

Function AlleBriefe()
' Dim q AS DAO.QueryDef
 Call do_Click_Vorbereit
 sql = "SELECT a.nachname +"",""+a.vorname Nachname, cstr(a.diabetestyp) vorname, b.pfad AS DokPfad, b.name DokName, b.art DokArt, b.* FROM `" + QMdbAkt + "`.`briefe` b LEFT JOIN `" + QMdbAkt + "`.Anamnesebogen a ON `briefe`.pat_id = a.pat_id WHERE `briefe`.name LIKE ""Brief an*"" AND NOT cstr(diabetestyp) IN (""1"", ""2"",""g"",""s"") ORDER BY zeitpunkt DESC"
 Call DtbCreateQueryDef("LaborDokumente eP", sql)
 DoCmd.OpenForm "LaborDokumente eP", acNormal, , "17=17"
 Forms![LaborDokumente ep].Caption = "Alle Briefe"
 ' irfan
End Function ' AlleBriefe()

#If zutesten Then
Function FormAusg()
 Dim i%, ct As Control, j%
 Open uVerz & "Anambog.txt" For Output As #293
 Call AnbogVar
 With Application.Forms(Anmnbi)
  For i = 0 To .Controls.COUNT - 1
   Set ct = .Controls(i)
'   Debug.Print ct.Properties.COUNT
   For j = 0 To ct.Properties.COUNT - 1
'    Debug.Print " " & ct.Properties(j).name
'    Debug.Print "  " & ct.Properties(j).Value
   Next j
  Next i
' For i = 1 To .Properties.Count
'  Debug.Print i & " " & .Properties(i).Name & " "; .Properties(i).Value
' Next i
 End With
 Close #293
End Function ' FormAusg()
#End If

#If False Then

DROP TABLE IF EXISTS `quelle`.`medarten`;
CREATE TABLE  `quelle`.`medarten` (
  `Medikament` VARCHAR(50) CHARACTER SET latin1 COLLATE latin1_german2_ci DEFAULT NULL,
  `Langname` VARCHAR(150) CHARACTER SET latin1 COLLATE latin1_german2_ci DEFAULT NULL COMMENT 'Beispiel-Langname',
  `Pat_ID` int(10) DEFAULT NULL COMMENT 'Beispiel-PatID',
  `Anzahl` int(10) DEFAULT NULL COMMENT 'Anzahl der Vorkommen',
  `Glib` bit(1) DEFAULT NULL COMMENT 'Glibenclamid',
  `Metf` bit(1) DEFAULT NULL COMMENT 'Metformin',
  `GlucI` bit(1) DEFAULT NULL COMMENT 'Glucosidase-Inhibitoren',
  `SHGlin` bit(1) DEFAULT NULL COMMENT 'andere Sulfonylharnstoffe oder Glinide',
  `Glit` bit(1) DEFAULT NULL COMMENT 'Glitanzone',
  `Ins` bit(1) DEFAULT NULL COMMENT 'Insulin',
  `Anal` bit(1) DEFAULT NULL COMMENT 'Insulin-Analoga',
  `InsArt` VARCHAR(1) CHARACTER SET latin1 COLLATE latin1_german2_ci DEFAULT NULL COMMENT '1= schnell, 2 = langsam, 3 = Misch',
  `HMG` bit(1) DEFAULT NULL COMMENT 'HMG-CoA-Reduktase-Inhibitoren',
  `Hypt` bit(1) DEFAULT NULL COMMENT 'Hypertonie-Mittel',
  `Thro` bit(1) DEFAULT NULL COMMENT 'Thrombozyten-Hemmer',
  `Antib` bit(1) DEFAULT NULL COMMENT 'Antibiotika',
  `and` bit(1) DEFAULT NULL COMMENT 'andere',
  `hinzugefügt` DATETIME DEFAULT NULL,
  `Tstr` bit(1) DEFAULT NULL COMMENT 'Teststreifen',
  `Puzu` bit(1) DEFAULT NULL COMMENT 'Pumpenzubehör',
  `VMat` bit(1) DEFAULT NULL COMMENT 'Verbandsmaterial',
  `PenN` bit(1) DEFAULT NULL COMMENT 'Pennadeln',
  `Neurp` bit(1) DEFAULT NULL COMMENT 'Neuropathie-Behandlungsmittel',
  `AutNP` bit(1) DEFAULT NULL COMMENT 'Autonome Neuropathie',
  `Fetts` bit(1) DEFAULT NULL COMMENT 'Fibrate, Ezetrol, Niaspan u.a.',
  `Hsre` bit(1) DEFAULT NULL COMMENT 'Hyperuriämie-Mittel',
  `AntiMyk` bit(1) DEFAULT NULL COMMENT 'Antimykotika',
  `Glauk` bit(1) DEFAULT NULL COMMENT 'Glaukom',
  `COLD` bit(1) DEFAULT NULL COMMENT 'COLD und Asthma',
  `Pros` bit(1) DEFAULT NULL COMMENT 'Prostatahypertrophie',
  `Urä` bit(1) DEFAULT NULL COMMENT 'Urämie-spezifische',
  `HyThy` bit(1) DEFAULT NULL COMMENT 'thyreostatische Mittel',
  `Ostp` bit(1) DEFAULT NULL COMMENT 'Osteoporosemittel',
  `KHK` bit(1) DEFAULT NULL COMMENT 'KHK-spezifisch',
  `HerzI` bit(1) DEFAULT NULL COMMENT 'Herzinsuffizienz-spezifische',
  `Stru` bit(1) DEFAULT NULL COMMENT 'Struma- und Hypothyreosemittel',
  `AVK` bit(1) DEFAULT NULL COMMENT 'AVK-Mittel',
  `PanI` bit(1) DEFAULT NULL COMMENT 'Pankreasinsuffizienz',
  `Vari` bit(1) DEFAULT NULL COMMENT 'Varikosemittel',
  `Östr` bit(1) DEFAULT NULL COMMENT 'Östrogene, Gestagene usw.',
  `AntiDep` bit(1) DEFAULT NULL COMMENT 'Antidepressiva',
  `AntiDem` bit(1) DEFAULT NULL COMMENT 'Antidementika',
  `AntiEp` bit(1) DEFAULT NULL COMMENT 'Antiepileptika',
  `Park` bit(1) DEFAULT NULL COMMENT 'Parkinson-Medikament',
  `AntiPern` bit(1) DEFAULT NULL COMMENT 'Antiperniziosa',
  `Appet` bit(1) DEFAULT NULL COMMENT 'Appetitzügler',
  `Anäm` bit(1) DEFAULT NULL COMMENT 'Anämiebehandlungsmittel',
  `Antiherp` bit(1) DEFAULT NULL COMMENT 'Anti-Herpes-Mittel',
  `NSAR` bit(1) DEFAULT NULL COMMENT 'NSAR',
  `Antikoag` bit(1) DEFAULT NULL COMMENT 'Antikoagulatien',
  `Betabl` bit(1) DEFAULT NULL COMMENT 'Betablocker',
  `ACEH` bit(1) DEFAULT NULL COMMENT 'ACE-Hemmer',
  `AT1` bit(1) DEFAULT NULL COMMENT 'AT-1-Blocker',
  `CalcA` bit(1) DEFAULT NULL COMMENT 'Calcium-Antagonist',
  `Diur` bit(1) DEFAULT NULL COMMENT 'Diuretikum',
  `falsch` bit(1) DEFAULT NULL COMMENT 'falsch geschrieben oder nicht erkennbar',
  UNIQUE KEY `Medikament` (`Medikament`),
  KEY `pat_id` (`Pat_ID`),
  CONSTRAINT `NamenMedArten_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;


DROP TABLE IF EXISTS `quelle`.`medplan`;
CREATE TABLE  `quelle`.`medplan` (
  `FID` int(10) DEFAULT NULL COMMENT 'Fall-Bezug',
  `Pat_ID` int(10) DEFAULT NULL COMMENT '3000',
  `MPNr` int(10) DEFAULT NULL COMMENT 'Ordnungsziffer für Medikamentenplan',
  `ZeitPunkt` DATETIME DEFAULT NULL COMMENT 'Zeitpunkt, der Speicherung im Turbomed',
  `Datum` DATETIME DEFAULT NULL COMMENT 'Zeitpunkt aus dem Kopf des Medikamentenplans',
  `Medikament` VARCHAR(75) COLLATE latin1_german2_ci DEFAULT NULL,
  `MedAnfang` VARCHAR(35) COLLATE latin1_german2_ci DEFAULT NULL,
  `FeldNr` SMALLINT(6) DEFAULT NULL,
  `mo` VARCHAR(10) COLLATE latin1_german2_ci DEFAULT NULL,
  `mi` VARCHAR(10) COLLATE latin1_german2_ci DEFAULT NULL,
  `nm` VARCHAR(10) COLLATE latin1_german2_ci DEFAULT NULL,
  `ab` VARCHAR(10) COLLATE latin1_german2_ci DEFAULT NULL,
  `zn` VARCHAR(10) COLLATE latin1_german2_ci DEFAULT NULL,
  `bBed` bit(1) DEFAULT NULL,
  `Bemerkung` longtext COLLATE latin1_german2_ci,
  `AbsPos` int(10) DEFAULT NULL COMMENT 'Zeile in der BDT-Datei',
  `AktZeit` DATETIME DEFAULT NULL COMMENT 'Aktualisierungszeit',
  `StByte` int(10) DEFAULT NULL COMMENT 'Ordnungsnummer der Datenübertragung',
  KEY `Auswahl` (`Pat_ID`,`ZeitPunkt`,`FeldNr`),
  KEY `FälleMedPlan` (`FID`),
  KEY `FID` (`FID`),
  KEY `MedPlanMedikament` (`Medikament`),
  KEY `NamenMedPlan` (`Pat_ID`),
  KEY `MedArtenMedPlan_AccRel` (`MedAnfang`),
  CONSTRAINT `FälleMedPlan_AccRel` FOREIGN KEY (`FID`) REFERENCES `faelle` (`FID`),
  CONSTRAINT `MedArtenMedPlan_AccRel` FOREIGN KEY (`MedAnfang`) REFERENCES `medarten` (`Medikament`),
  CONSTRAINT `NamenMedPlan_AccRel` FOREIGN KEY (`Pat_ID`) REFERENCES `namen` (`Pat_ID`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COLLATE=latin1_german2_ci;


ALTER TABLE medplan drop FOREIGN KEY MedArtenMedPlan_AccRel;
ALTER TABLE medplan MODIFY COLUMN `MedAnfang` VARCHAR(35) CHARACTER SET latin1 COLLATE latin1_german2_ci DEFAULT NULL;
DELETE FROM `medarten`;
ALTER TABLE `medarten` modify column Medikament VARCHAR(50) CHARACTER SET latin1 COLLATE latin1_german2_ci DEFAULT NULL;
INSERT INTO `medarten` (medikament) SELECT DISTINCT medanfang FROM medplan m;
ALTER TABLE  medplan ADD CONSTRAINT `MedArtenMedPlan_AccRel` FOREIGN KEY (`MedAnfang`) REFERENCES `medarten` (`Medikament`);

#End If

#If zutesten Then
Function asctest()
Dim i%, j%
For i = 0 To 255
 Debug.Print i, Chr$(i)
Next i
End Function ' asctest
#End If

'Function zeigviews()
' Dim rv As New ADODB.Recordset, rCn As New ADODB.Connection, i%, stri$
'' Call rCn.Open(CStrAcc & StACCDB)
' CStrMy = "DRIVER={" & ODBCStr & "};server=" & LiName & ";user=praxis;pwd=...;database="
' Call rCn.Open(CStrMy & "quelle")
'' SET rV = rCn.OpenSchema(adSchemaViews)
' Set rv = rCn.OpenSchema(adSchemaTables)
' Do While Not rv.EOF
'  If rv!table_type = "VIEW" Then
'   stri = vNS
'   For i = 0 To rv.Fields.COUNT - 1
'    stri = stri & " " & rv.Fields(i).name & ":" & rv.Fields(i)
'   Next i
'   Debug.Print stri
'  End If
'  rv.Move 1
' Loop
'End Function ' zeigviews

'Function btest()
' Dim rv As ADODB.Recordset, rCn As New ADODB.Connection, i%, stri$, rAF&
'' Call rCn.Open(CStrAcc & StACCDB)
' CStrMy = "DRIVER={" & ODBCStr & "};server=" & LiName & ";user=praxis;pwd=...;database="
' Call rCn.Open(CStrMy & "quelle")
' myFrag rv, "SELECT tkz,pat_id FROM `anamnesebogen` WHERE pat_id=2", adOpenStatic, rCn, adLockOptimistic
' rv!Tkz = 1
' If rv.EditMode = adEditInProgress Then rv.Update
' Debug.Print rv!Tkz
' rv!Tkz = 0
' If rv.EditMode = adEditInProgress Then rv.Update
' Debug.Print rv!Tkz
' rv!Tkz = 1
' If rv.EditMode = adEditInProgress Then rv.Update
' Debug.Print rv!Tkz
' myEFrag "UPDATE `anamnesebogen` SET tkz = 1 WHERE pat_id = 2", rAF, rCn
' myFrag rv, "SELECT tkz,pat_id FROM `anamnesebogen` WHERE pat_id=2", adOpenStatic, rCn, adLockOptimistic
' Debug.Print "rAf=", rAF, rv!Tkz
' myEFrag "UPDATE `anamnesebogen` SET tkz = 0 WHERE pat_id = 2", rAF, rCn
' myFrag rv, "SELECT tkz,pat_id FROM `anamnesebogen` WHERE pat_id=2", adOpenStatic, rCn, adLockOptimistic
' Debug.Print "rAf=", rAF, rv!Tkz
' myEFrag "UPDATE `anamnesebogen` SET tkz = -1 WHERE pat_id = 2", rAF, rCn
' myFrag rv, "SELECT tkz,pat_id FROM `anamnesebogen` WHERE pat_id=2", adOpenStatic, rCn, adLockOptimistic
' Debug.Print "rAf=", rAF, rv!Tkz
'End Function ' btest

Function dovCommandB_Click(Index%)
'
End Function ' dovCommandB_Click(Index%)


Function do_abgehakt_Click(frm As Form) ' für Labordokumente eP
 Static Pat_ID&, cR
 On Error GoTo fehler
 Dim rs As New ADODB.Recordset
 myFrag rs, frm.anaRS.source
 If Not rs.BOF Then
'  myEFrag("UPDATE ")
 Else
'
 End If
' rDokab.Seek "=", frm.DokPfad
' IF Not rDokab.NoMatch THEN
'  rDokab.Edit
' Else
'  rDokab.AddNew
'  rDokab!DokPfad = frm.DokPfad
' END IF
' IF rDokab!abgehakt THEN
'  rDokab!abgehakt = False
' Else
'  rDokab!abgehakt = True
' END IF
' rDokab!AktZeit = Now
' rDokab.Update
' Pat_id = frm.Pat_id
' cR = frm.CurrentRecord
  frm.Requery
 On Error Resume Next
' DoCmd.GoToRecord acDataForm, frm.Name, acGoTo, cR
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in do_abgehakt_Click/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_abgehakt_Click(frm AS Form)

' Labordokumente eP
Function doDatensatzPosition$(frm As Form)
  Dim rs As New ADODB.Recordset
  If Not IsNull(frm.anaRS!DokPfad) Then 'vTextB(6)) THEN ' DokPfad
    myFrag rs, frm.anaRS.source
    rs.Find "dokpfad = '" & frm.anaRS!DokPfad & "'"
    If Not rs.EOF Then
     doDatensatzPosition = CStr(rs.AbsolutePosition + 1)
    End If
'    SET rs = frm.RecordsetClone
'    rs.FindFirst ("LaborDokumente.Dokpfad=" & vns & frm!DokPfad & vns)
'    DatensatzPosition = CStr(rs.AbsolutePosition + 1)
'    rs.Close
  End If
End Function ' doDatensatzPosition

Function doAktAbgehakt(frm As Form)
 frm.anaRS.Find "dokpfad = '" & frm.DokPfad & "'"
 If frm.anaRS.EOF Then
  doAktAbgehakt = False
 Else
  doAktAbgehakt = frm.anaRS!abgehakt
 End If
' rDokab.Seek "=", frm.DokPfad
' IF rDokab.NoMatch THEN
'  AktAbgehakt = False
' Else
'  AktAbgehakt = rDokab!abgehakt
' END IF
End Function ' doAktAbgehakt(frm As Form)

Function do_obAbgehakt(frm As Form)
 frm.anaRS.Find "dokpfad = '" & frm.DokPfad & "'"
 If frm.anaRS.EOF Then
  do_obAbgehakt = vNS
 Else
  If frm.anaRS!abgehakt <> 0 Then
   do_obAbgehakt = "ja"
  Else
   do_obAbgehakt = vNS
  End If
  do_obAbgehakt = frm.anaRS!abgehakt
 End If
' rDokab.Seek "=", frm.DokPfad
' IF rDokab.NoMatch THEN
'  obAbgehakt = vns
' Else
'  IF rDokab!abgehakt THEN
'   obAbgehakt = "ja"
'  Else
'   obAbgehakt = vns
'  END IF
' END IF
End Function

Function doAltName(frm As Form)
 Static AltN$, NeuN$, altMerk$
 NeuN = frm.Nachname + frm.Vorname
 If NeuN <> AltN Then
  If altMerk = "neu" Then
   altMerk = vNS
  Else
   altMerk = "neu"
  End If
 End If
 doAltName = altMerk
 AltN = NeuN
End Function ' doAltName

Function doForm_Load(frm As Form)
  Dim db As New Connection
  Dim rs As ADODB.Recordset
  Dim oText As TextBox
  Dim oCheck As CheckBox, ctl
  Dim Pat_ID&, rAf&
  On Error GoTo fehler
  db.CursorLocation = adUseClient
  db.Open Lese.dbv.CnStr
  Set frm.anaRS = New Recordset
  
  If frm.name Like "AnBog*" Then
   syscmd 4, "Lade Datensätze von der Datenbank ..."
   frm.obStumm = True
   If LVobMySQL Then
'    frm.anaRS.Open "SELECT *,IF(größe=0,'',gewicht/größe/größe*IF(größe>3,10000,1)) AS bmi, CONCAT(nachname,' ', vorname) AS gesname  FROM `anamnesebogen` WHERE pat_id IN (SELECT pat_id FROM `aktfv`) ORDER BY pat_id", db, adOpenStatic, adLockOptimistic
'    frm.anaRS.Open "SELECT a.*, CONCAT(nachname,' ', vorname) AS gesname FROM `anamnesebogen` WHERE exists (SELECT * FROM `aktfv` WHERE pat_id = `anamnesebogen`.pat_id) ORDER BY pat_id DESC", db, adOpenDynamic, adLockOptimistic
    Dim sqlspl$(), orderby$
    SplitNeu frm.PidRange, "ORDER BY", sqlspl
    If UBound(sqlspl) > 0 Then orderby = sqlspl(1) Else orderby = "pat_id DESC"
'    frm.anBogCS = "SELECT MAX(f.fid) fid, `anamnesebogen`.* FROM `anamnesebogen` LEFT JOIN `faelle` f ON `anamnesebogen`.pat_id = f.pat_id WHERE exists (SELECT pat_id FROM (" & sqlspl(0) & ") i WHERE pat_id = `anamnesebogen`.pat_id) GROUP BY `anamnesebogen`.pat_id ORDER BY " & orderby
'    frm.anBogCS = "SELECT pat_id, nachname, vorname,`Vibration IK`,`Vibration Großzehe`,`Ameisenlaufen`,`Kraft Zehenheber`,`Kraft Zehenbeuger`,`Kraft Knie`,ASR,PSR,`Oberflächensensibilität`,`Monofilamenttest`,`Kalt-Warm`,`Liphypertrophien Abdomen`,`Liphypertrophien Arme`,`Liphypertrophien Beine`,`Puls Leiste`,`Puls Kniekehle`,`Puls Atp`,`Puls Adp`,bmi,Bluthochdruck,`Augensp Befund`,`Weitere Anamnese`,`Sehminderung unbehebbar`,`Netzhaut gelasert`,`Diabet Nierenschaden`,Dialyse,`andere Nierenerkrankung`,`Angina pectoris`,`Herzinfarkt`,`PTCA oder Stent`,`Bypass kardial`,`Herzkrankheit Beschreibung`,Hirndurchblutungsstörung,Schlaganfall,Beindurchblutungsstörung,Schaufensterkrankheit,`bypaß peripher`,Geschwür,Amputation,`pavk beschreibung`,Diabetestyp,Insulinpumpe,tkz,oban1eing,obanaeing,oban2eing,obcheck,obbzausgew,obosaufgek,obpodaufgek,obmblausgeh,obmednetz,prim,gebdat,`diabetes seit`,`tabletten seit`,`insulin seit`,`grund für vorstellung`,familienanamnese,größe,gewicht,`letztes HbA1c`,Anrede,ther1," & _
    "`BZMessungen selbst`,Herzkrankheit,Sexualstörung,Verformungen,Bewegungseinschränkungen,`Entleerungsstörungen Magen`,`Entleerungsstörungen Harnblase`,Ulcera,Herzschwäche,Lunge,ws,sd,`sehminderung unbehebbar`,`Folgeerkrankungen Haut`,Herz,Zähne,Hyperkeratosen,BeinödVen,Beinbefund,Tendenz,`Diabetesmedikament 1`,`Diabetesmedikament 1 Menge`,`Diabetesmedikament 2`,`Diabetesmedikament 2 Menge`,`Diabetesmedikament 3`,`Diabetesmedikament 3 Menge`,`Diabetesmedikament 4`,`Diabetesmedikament 4 Menge`,`Insulinpumpe seit`,`Insulinpumpe Marke`,`Broteinheiten gesamt`,`Broteinheiten früh`,`Broteinheiten zm früh`,`Broteinheiten mittags`,`Broteinheiten nachmittags`,`Broteinheiten abends`,`Broteinheiten nachts`,`Essenszeit früh` " & _
    " FROM `anamnesebogen` WHERE exists (SELECT pat_id FROM (" & sqlspl(0) & ") i WHERE pat_id = `anamnesebogen`.pat_id) GROUP BY `anamnesebogen`.pat_id ORDER BY " & orderby
    'frm.anBogCS = "SELECT * FROM `anakt` WHERE exists (SELECT pat_id FROM (" & sqlspl(0) & ") i " & _
    "WHERE pat_id = `anakt`.pat_id) ORDER BY " & orderby
'    frm.anBogCS = "SELECT a.* FROM `anakt` a LEFT JOIN aktfv f USING (pat_id) WHERE NOT ISNULL(f.pat_id) ORDER BY vorgestellt DESC"
    frm.anBogCS = "SELECT * FROM `anakt` WHERE pat_id IN (" & frm.PidRange & ") ORDER BY vorgestellt DESC"
'    frm.anBogCS = "SELECT * FROM `anamnesebogen` ORDER BY pat_id DESC"
    db.CursorLocation = adUseClient
    frm.anaRS.CursorLocation = adUseClient
    frm.anaRS.CursorType = adOpenDynamic
    db.DefaultDatabase = "quelle"
    myEFrag "UPDATE `anamnesebogen` SET bmi = IF(größe=0,0,IF(gewicht>0,gewicht,-gewicht)/größe/größe*IF(größe>3,10000,1)) WHERE bmi <> IF(größe=0,0,gewicht/größe/größe*IF(größe>3,10000,1))", rAf, db ' 13.9.09
'    frm.anaRS.Open frm.anBogCS, db, adOpenDynamic, adLockOptimistic
    Set frm.anaRS = myFrag(rs, frm.anBogCS, adOpenDynamic, db)
   Else
    db.CursorLocation = adUseClient
'    frm.anaRS.Open "SELECT *,IIF(größe=0,'',gewicht/größe/größe*iIF(größe>3,10000,1)) bmi, GesName(pat_id) gesname  FROM `anamnesebogen` ORDER BY pat_id DESC", db, adOpenDynamic, adLockOptimistic
    Set frm.anaRS = myFrag(rs, "SELECT *,IIF(größe=0,'',gewicht/größe/größe*iIF(größe>3,10000,1)) bmi, GesName(pat_id) gesname  FROM `anamnesebogen` ORDER BY pat_id DESC", adOpenDynamic, db)
   End If
   frm.obStumm = False
   On Error Resume Next
   Pat_ID = fWertLesen(HCU, RegWurzel & App.EXEName, "pat_id")
   On Error GoTo fehler
   If Not myEFrag("SELECT 0 FROM (" & frm.anBogCS & ") i WHERE pat_id=" & Pat_ID).BOF Then
    If Pat_ID = 0 Then frm.anaRS.MoveFirst Else frm.anaRS.Find "pat_id = " & Pat_ID
   Else ' dann muss do_Form_Current nochmal ohne stumm nachgeholt werden
    frm.anaRS.Move 0
   End If
  'Kontrollkästchen an Datenprovider binden
   For Each oCheck In frm.vCheckb
    Set oCheck.DataSource = frm.anaRS
   Next
'   For Each ctl In frm.Controls
'    Err.Clear
'    ON Error Resume Next
'    IF ctl.DataSource <> frm.anaRS THEN
'     SET ctl.DataSource = frm.anaRS
'    END IF
'   Next ctl
  ElseIf frm.name = "Medarten" Then
   syscmd 4, "Lade Datensätze von der Datenbank ..."
   frm.obStumm = True
   frm.anaRS.Open "SELECT * FROM `medarten` ORDER BY (glib<>0 OR metf<>0 OR gluci<>0 OR shglin<>0 OR glit<>0 OR dpp4<>0 OR glp1<>0 OR sglt2<>0 OR sonstad<>0  OR ins<>0 OR anal<>0 OR insart<>0 OR hmg<>0 OR hypt<>0 OR thro<>0 OR antib<>0 OR `and`<>0 OR tstr<>0 OR puzu<>0 OR vmat<>0 OR penn<>0 OR neurp<>0 OR autnp<>0 OR fetts<>0 OR hsre<>0 OR antimyk<>0 OR glauk<>0 OR cold<>0 OR pros<>0 OR `urä`<>0 OR hythy<>0 OR ostp<>0 OR khk<>0 OR herzi<>0 OR stru<>0 OR avk<>0 OR pani<>0 OR vari<>0 OR `östr`<>0 OR antidep<>0 OR antidem<>0 OR antiep<>0 OR park<>0 OR antipern<>0 OR appet<>0 OR anäm<>0 OR antiherp<>0 OR nsar<>0 OR antikoag<>0 OR betabl<>0 OR aceh<>0 OR at1<>0 OR calca<>0 OR diur<>0 OR falsch<>0), id DESC", db, adOpenStatic, adLockOptimistic
' das folgende funktionierte auf szn4, nicht aber auf anmmw 30.12.22
'   myFrag frm.anaRS, "SELECT * FROM `medarten` ORDER BY (glib<>0 OR metf<>0 OR gluci<>0 OR shglin<>0 OR glit<>0 OR dpp4<>0 OR glp1<>0 OR sglt2<>0 OR sonstad<>0  OR ins<>0 OR anal<>0 OR insart<>0 OR hmg<>0 OR hypt<>0 OR thro<>0 OR antib<>0 OR `and`<>0 OR tstr<>0 OR puzu<>0 OR vmat<>0 OR penn<>0 OR neurp<>0 OR autnp<>0 OR fetts<>0 OR hsre<>0 OR antimyk<>0 OR glauk<>0 OR cold<>0 OR pros<>0 OR `urä`<>0 OR hythy<>0 OR ostp<>0 OR khk<>0 OR herzi<>0 OR stru<>0 OR avk<>0 OR pani<>0 OR vari<>0 OR `östr`<>0 OR antidep<>0 OR antidem<>0 OR antiep<>0 OR park<>0 OR antipern<>0 OR appet<>0 OR anäm<>0 OR antiherp<>0 OR nsar<>0 OR antikoag<>0 OR betabl<>0 OR aceh<>0 OR at1<>0 OR calca<>0 OR diur<>0 OR falsch<>0), id DESC", adOpenStatic, db, adLockOptimistic
  'Kontrollkästchen an Datenprovider binden
   For Each oCheck In frm.vCheckb
    Set oCheck.DataSource = frm.anaRS
   Next
  End If
  'Textfelder an Datenprovider binden
   For Each oText In frm.vTextB
    Set oText.DataSource = frm.anaRS
   Next
  frm.mbDataChanged = False
  Lese.Hide
 'Me.Controls!Diagnosen.ControlSource = "=replace$(replace$(recordset!Diagnosen,vbverticaltab,vbcr+vblf),vbtab,"" "")"
  Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in doForm_Load/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doFormLoad

Function doKeyDown(frm As Form, KeyCode%, Shift%)
 If frm.name Like "AnBog*" Then
   Select Case KeyCode
    Case 27
     Call AnBogUnload(frm)
   End Select
 ElseIf frm.name Like "Medarten*" Then
   Select Case KeyCode
    Case 27
     Unload frm
   End Select
 End If
End Function ' FUNCTION doKeyDown(frm AS Form, keyCode%, Shift%)

Function AnBogUnload(frm As AnBog)
  ' frm.anaRS!Pat_id
 If Not frm.anaRS.EOF And Not frm.anaRS.BOF Then _
  Call fDWSpei(HCU, RegWurzel & App.EXEName, "pat_id", Pat_ID)       ' für MySQL
 Unload frm
End Function ' FUNCTION AnBogUnload(frm AS AnBog)

' in GesDiagExp
Function diagexpHerricht()
 Dim rAf&
 On Error GoTo fehler
 myEFrag ("DELETE FROM `diagnosenexport`")
 myEFrag "INSERT INTO `diagnosenexport` (id,name,pat_id,icd,diagnose,status,protokoll,nurquart,zeitpunkt) SELECT id,name,pat_id,icd,diagnose,status,protokoll,nurquart,Zeitpunkt FROM `fuerdiagexp` WHERE icd <> ''", rAf
 Exit Function
fehler:
Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in diagexpHerricht/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' diagexpHerricht()

' DiagnosenExportierenTest_Click
Function GesDiagExp(Optional obTest%) ' Im Menü unter "Diagnosen exportieren"
 On Error GoTo fehler
 Call diagexpHerricht

' ON Error Resume Next
' Dim test$
' test = Forms!Anamnesebogen.Name
' IF Err.Number = 0 THEN
'  ON Error GoTo fehler
'  Call do_Diagnosen_Reset(Forms!Anamnesebogen) ' muß offen sein
' Else
'  ON Error GoTo fehler
'  Call do_Diagnosen_Reset
' END IF
 Call doDiagnosenexport(obTest:=obTest)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GesDiagExp/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DiagExpGesamt()

Function doRückgängig()
 Dim rAf&
 InsKorr DBCn, "INSERT INTO `fuerdiagexp`(name,pat_id,icd,diagnose,zeitpunkt) SELECT CONCAT(nachname, ' ', vorname) AS Name, e.pat_id, icd, diagnose, übertragen FROM `diagnosen exportiert` e LEFT JOIN `namen` USING (pat_id) WHERE übertragen = (SELECT MAX(übertragen) FROM `diagnosen exportiert`)", rAf
 syscmd 4, rAf & " Diagnosen erneut vorbereitet"
End Function ' doRückgängig()

#If Not thaalt Then
' #If zutesten Then
Function testthap(pids$)
' IF DBCn.State <> 0 THEN DBCn.Close
' SET DBCn = Nothing
 Lese.ProgStart
' 22.10.22: führt bei Aufruf über Ado zumindest bis zur Mariadb-Version 10.9 immer wieder zum Server-Crash, s.ähnliche Bug-Hinweise früherer Versionen
'#Const mitfensterther = True
#If mitfensterther Then
 rufauf "ssh", "root@" & LiName & " mysql --defaults-extra-file=~/.mysqlpwd quelle -e'CALL fuellThaP(" & pids & ")'", 2, "c:\windows\system32\openssh\", -1, 0
#Else
 Call TheraErmitt(pids)
#End If
' myEFrag "CALL fuellThaP(" & pids & ")"
 Debug.Print myEFrag("SELECT COUNT(0) zl FROM therarten WHERE pat_id IN (" & pids & ")").Fields(0)
 Lese.ProgEnde
 Debug.Print "Fertig"
End Function ' testthap(pids$)
' #End If
#End If

'Function testnull()
' Lese.ProgStart
' Dim q As New ADODB.Recordset
' q.Open "SELECT DISTINCT nurquart FROM `diagnosenexport` ORDER BY nurquart DESC", DBCn, adOpenDynamic, adLockOptimistic
' Do While Not q.EOF
'   IF NOT ISNULL(q!nurquart) AND q!nurquart = 1 THEN ' => stimmt
'    Debug.Print q!nurquart & " Quartal"
'   END IF
'   IF ISNULL(q!nurquart) OR q!nurquart = 0 THEN
'    Debug.Print q!nurquart & " Dauer"
'   END IF
'   q.MoveNext
' Loop
'End Function
' #end if ' zutesten

Function doDiagnosenexport(Optional obTest%)
 Dim Lanr&
 On Error GoTo fehler
 Dim Quartal$, erg$, dzahl&, BDT As New BDTSchreib
 
' Dim Q AS DAO.Recordset, rFa AS DAO.Recordset, n AS DAO.Recordset, ü AS DAO.Recordset
 Dim q As New ADODB.Recordset, rFa As New ADODB.Recordset, n As New ADODB.Recordset, Ü As New ADODB.Recordset
' Dim rDT As New ADODB.Recordset
' Dim z$, ZDat$ ', DDat As Date, rohdat$
 Dim aktdat As Date
 Dim Pat_ID&
' Const quartal$ = "22005"
 On Error GoTo fehler
 Call syscmd(acSysCmdSetStatus, "Diagnosen exportieren ...")
' Call ImportFolderHerricht
' Do While Not IsDate(rohdat)
'  rohdat = InputBox("Für welches Datum sollen die Diagnosen eingetragen werden?")
' Loop
' DDat = CDate(rohdat)

' ZDat = "DIAG " + Format$(üzpt, "dd/mm/yy HH.MM") + ".BDT"
' z = hVerz & ZDat
' Open z For Output AS #327
' Close #327
' erg = Dir(hVerz & ZDat)
' IF LenB(erg) = 0 THEN
'  MsgBox "Fehler beim Erstellen der Exportdatei '" & hVerz & ZDat & "'"
'  Exit Function
' END IF
' SET Q = Dtb.OpenRecordset("DiagnosenExport", dbOpenDynaset)
' q.Open "SELECT * FROM `diagnosenexport` WHERE (ISNULL(`status`) OR `status` <> 'übertragen') AND NOT ISNULL(pat_id) AND pat_id <> 0 ORDER BY pat_id DESC, nurquart DESC, zeitpunkt DESC", DBCn, adOpenDynamic, adLockOptimistic
 myFrag q, "SELECT * FROM `diagnosenexport` WHERE (ISNULL(`status`) OR `status` <> 'übertragen') AND NOT ISNULL(pat_id) AND pat_id <> 0 ORDER BY pat_id DESC, nurquart DESC, zeitpunkt DESC"
 If q Is Nothing Then
  MsgBox "Fehler beim  Öffnen der Tabelle `diagnosenexport`"
  Exit Function
 ElseIf q.BOF Then
  MsgBox "Keine Diagnosen zum Exportieren!"
  Exit Function
 End If
 
' SET ü = Dtb.OpenRecordset("Diagnosen exportiert", dbOpenTable)
' Ü.Open "SELECT * FROM `diagnosen exportiert`", DBCn, adOpenDynamic, adLockOptimistic
 myFrag Ü, "SELECT * FROM `diagnosen exportiert`"
 If Ü Is Nothing Then
  MsgBox "Fehler beim  Öffnen der Tabelle `diagnosen exportiert`"
  Exit Function
 End If
 
' Open z For Append AS #327
' SET rFa = TabÖff("faelle", "Auswahl")
' SET n = TabÖff("Namen", "Pat_ID")
 
 If Not q.BOF Then
  If Not BDT.Start(hVerz, "DIAG") Then ' Datei öffnen
   Exit Function
  End If
  Call BDT.ImportFolderHerricht
  Call BDT.BDTKopf
'  SET rDT = TabÖff("Diagnosen", "DiagSuch")
'  SET rDT = Nothing
'  rDT.Open "SELECT * FROM diagnosen", DBCn, adOpenDynamic, adLockOptimistic
  Pat_ID = -1
  Dim neuDauer%, obDauer%
  Dim nurquart%, ICD$, Zeitpunkt As Date, Diagnose$, DiagText$, id&, name$
  Do While Not q.EOF
   nurquart = q!nurquart
   ICD = q!ICD
   Zeitpunkt = q!Zeitpunkt
   Diagnose = q!Diagnose
'   DiagText = q!DiagText
   name = q!name
   id = q!id
   If q!Pat_ID <> Pat_ID Then
    Pat_ID = q!Pat_ID ' Hier kommt er nur einmal pro Patient vorbei
    neuDauer = 0
    If IsNull(Zeitpunkt) Then aktdat = lFDat(Pat_ID) Else aktdat = Zeitpunkt
    Quartal = ZQuart(aktdat)
    Set rFa = Nothing
'    rFa.Open "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_id & " AND quartal = '" & Quartal & "' ORDER BY bhfb DESC", DBCn, adOpenDynamic, adLockOptimistic
    myFrag rFa, "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_ID & " AND quartal = '" & Quartal & "' ORDER BY bhfb DESC"
    If rFa.BOF Then
     Set rFa = Nothing
'     rFa.Open "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_id & " ORDER BY bhfb DESC", DBCn, adOpenDynamic, adLockOptimistic
     myFrag rFa, "SELECT * FROM `faelle` f LEFT JOIN `lanrpraxis` l ON f.lanrid = l.id WHERE pat_id = " & Pat_ID & " ORDER BY bhfb DESC"
    End If
    If Not rFa.BOF Then
     Call FallExport(BDT, Pat_ID, aktdat, Lanr)
    End If
   End If ' Pat_id <> Pat_id THEN
   If Not rFa.BOF Then
    If Not IsNull(nurquart) Then ' 14.6.21
     obDauer = Not nurquart
    Else
     obDauer = 0
    End If
    If obDauer <> 0 And neuDauer = 0 Then
     neuDauer = 1
     BDT.Satzart "6200"
     BDT.PatID Pat_ID
     BDT.Satzart "6100"
     BDT.PatID Pat_ID
    End If ' q!obDauer <> 0 AND neuDauer = 0 THEN
        
        BDT.DAdd IIf(obDauer = 0, "5999", "3649"), aktdat
        BDT.TAdd "6201", aktdat
        BDT.SAdd "6203", "TM#?##"
        BDT.SAdd "3635", "TM#" & rFa!Lanr
        BDT.SAdd "3636", "TM#" & BSNR
        Dim DiagSi$ ', DiagText$
'        ICD = ICD
        DiagSi = "G"
        DiagText = Diagnose
        If Not IsNull(ICD) Then
         If ICD <> vNS Then
          If InStrB("VGZA", Right$(ICD, 1)) > 0 And Right$(ICD, 1) <> vNS Then
           DiagSi = Right$(ICD, 1)
           ICD = Left$(ICD, Len(ICD) - 1)
           Select Case DiagSi
            Case "V": DiagText = LTrim$(REPLACE$(DiagText, "V.a.", vNS))
            Case "Z": DiagText = LTrim$(REPLACE$(DiagText, "Z.n.", vNS))
            Case "A": DiagText = LTrim$(REPLACE$(DiagText, "Ausschluss ", vNS))
           End Select
          End If
         End If ' ICD <> vNS THEN
        End If ' NOT ISNULL(ICD) THEN
        BDT.SAdd IIf(obDauer = 0, "6000", "3650"), DiagText
        BDT.SAdd "6001", ICD
        BDT.SAdd "6003", DiagSi
'        BDT.SAdd "6010", "TM#Falsch"
        BDT.SAdd "6010", "TM#False"   ' 26.10.15, offenbar im Datensatz geändert
        BDT.SAdd "6011", "TM#?"
        dzahl = dzahl + 1
        
       If Not obTest Then
         Dim rAf&, rAfL&
         InsKorr DBCn, "INSERT INTO `diagnosen`(pat_id, ICD,diagdatum,diagsicherheit,diagtext,obdauer,aktzeit) VALUES(" & Pat_ID & ",'" & ICD & "'," & DatFor_k(aktdat) & ",'" & DiagSi & "','" & DiagText & "'," & IIf(obDauer = 0, 0, 1) & "," & DatFor_k(BDT.üzpt) & ")", rAf
         If rAf <> 1 Then
          MsgBox "Fehler beim Diagnoseneeinfügen für Pat. " & Pat_ID & vbCrLf & "ICD: " & ICD & vbCrLf & "Diagtext:" & DiagText & vbCrLf & "Datum: " & DatFor_k(aktdat) & rAf & " Datensätze eingefügt"
         End If
         If LenB(ICD) <> 0 And LenB(Diagnose) <> 0 Then
          Call myEFrag("UPDATE `diagnosenexport` SET status = '" & übertragen & "' WHERE id = " & id, rAf)
          If rAf <> 1 Then
           MsgBox "Fehler beim Statussetzen in `diagnosenexport` für ID: " & id & rAf & " Datensätze gesetzt"
          End If
          InsKorr DBCn, "INSERT INTO `diagnosen exportiert`(pat_id,datum,icd,diagnose,übertragen) VALUES(" & Pat_ID & "," & DatFor_k(aktdat) & ",'" & ICD & "','" & DiagText & "'," & DatFor_k(BDT.üzpt) & ")", rAf
          If rAf > 0 Then
           Call myEFrag("DELETE FROM `fuerdiagexp` WHERE id = " & id, rAfL)
           If rAfL <> 1 Then
            MsgBox "Fehler beim Löschen aus `fuerdiagexp` von " & Pat_ID & " (" & UmwfSQL(name) & ")" & vbCrLf & "ICD: " & ICD & vbCrLf & "Diagtext:" & DiagText & vbCrLf & "Datum: " & DatFor_k(aktdat) & vbCrLf & rAfL & " Datensätze gelöscht"
           End If
          End If
          If rAf <> 1 Then
           MsgBox "Fehler beim Eintragen in `diagnosen exportiert` von " & Pat_ID & vbCrLf & "ICD: " & ICD & vbCrLf & "Diagtext:" & DiagText & vbCrLf & "Datum: " & DatFor_k(aktdat) & rAf & " Datensätze eingetragen"
          End If
         End If ' icd <> vns AND Diagnose <> vns THEN
         Call dynDiag(CStr(Pat_ID)) ' 12.7.08
        End If ' obTEst
    
   End If ' Not rFa.BOF THEN
   q.MoveNext
  Loop ' While Not q.EOF
  End If ' not q.bof
  
'  Do While Not q.EOF
''   IF q!Pat_id = 77 THEN
'     IF ISNULL(q!Zeitpunkt) THEN aktdat = lFDat(q!pat_id) ELSE aktdat = q!Zeitpunkt
'     Quartal = ZQuart(aktdat)
'     IF q!pat_id <> altpat_id THEN
'      SET rFa = Nothing
'      rFa.Open "SELECT * FROM `faelle` WHERE pat_id = " & q!pat_id & " AND quartal = '" & Quartal & "' ORDER BY bhfb DESC", DBCn, adOpenDynamic, adLockOptimistic
''      rFa.Seek "=", Q!Pat_id, quartal
''      IF rFa.NoMatch THEN
'      IF rFa.BOF THEN
'       SET rFa = Nothing
'       rFa.Open "SELECT * FROM `faelle` WHERE pat_id = " & q!pat_id & " ORDER BY bhfb DESC", DBCn, adOpenDynamic, adLockOptimistic
''       rFa.Seek "=", Q!Pat_id
'      END IF
'      IF Not rFa.BOF THEN
'       Call FallExport(BDT, q!pat_id, aktdat, LANR)
'#If False THEN
''      IF Not rFa.NoMatch THEN
'       SET n = Nothing
'       n.Open "SELECT * FROM `namen` WHERE pat_id = " & q!pat_id, DBCn, adOpenDynamic, adLockOptimistic
''       n.Seek "=", Q!Pat_id
'       IF Not n.BOF THEN
''       IF Not n.NoMatch THEN
'        BDT.Satzart ("6200") ' IIF(rFa!SchGr = "90", "0190", "0102") ' Satzidentifikation (s.geslies)
'' bei 0101 entstehen bei zwei Aufrufen fehlerfrei zwei neue Kassenfaelle, jeder mit der Leistung
'' bei 190 entsteht ein neuer Privatfall, bei 6100 läuft alles ohne Fehler durch, aber keine Leistung steht drin
'' bei 6200 entsteht ein neuer Kassenfall
''        Print #327, "014810000845" ' Satzlänge, noch nicht entschlüsselt"
''     Dim op$
''     op = format$(3 + 4 + 4, "000") + "8000" + CStr(f!s8000)
''     Print #327, zsu(op)
''     op = format$(3 + 4 + 5, "000") + "8100" + CStr(f!s8100)
''     Print #327, zsu(op)
'        BDT.PatID q!pat_id
''        IF 1 = 0 THEN 'Auswirkung bisher nicht geprüft 31.7.05 (3x)
'        IF NOT ISNULL(n!NVorsatz) THEN
'         IF n!NVorsatz <> vNS THEN
'          BDT.NVorsatz n!NVorsatz
'         END IF
'        END IF
'        BDT.Nachname n!Nachname
'        BDT.Vorname n!Vorname
'        BDT.Geb n!GebDat
''        IF 1 = 0 THEN ' 26.6.08: Für Diagnosen nicht nötig
''         IF NOT ISNULL(n!Versichertennummer) THEN
''          op = Format$(3 + 4 + Len(n!Versichertennummer), "000") + "3105" + n!Versichertennummer
''          Print #327, ZSU(op)
''         END IF
'''     op = format$(3 + 4 + Len(n!Straße), "000") + "3107" + n!Straße
'''     Print #327, ZSU(op)
'''     op = format$(3 + 4 + Len(n!Plz), "000") + "3112" + n!Plz
'''     Print #327, ZSU(op)
'''     op = format$(3 + 4 + Len(n!Ort), "000") + "3113" + n!Ort
'''     Print #327, ZSU(op)
''         IF NOT ISNULL(n!KVKStatus) AND n!KVKStatus > 0 AND Trim$(n!KVKStatus) <> vns THEN
''          op = Format$(3 + 4 + Len(n!KVKStatus), "000") + "3108" + n!KVKStatus
''          Print #327, ZSU(op)
''         END IF
''         op = Format$(3 + 4 + 1, "000") + "3110" + IIF(n!Geschlecht = "m", "1", IIF(n!Geschlecht = "w", "2", vns))
''         Print #327, ZSU(op)
''         IF rFa!SchGr <> "90" THEN
''          op = Format$(3 + 4 + 5, "000") + "4101" + ZQuart(rFa!ausgst) 'quartal
''          Print #327, ZSU(op)
''          op = Format$(3 + 4 + 8, "000") + "4102" + Format$(rFa!ausgst, "ddmmyyyy")
''          Print #327, ZSU(op)
''          op = Format$(3 + 4 + Len(rFa!VKNr), "000") + "4104" + rFa!VKNr
''          Print #327, ZSU(op)
''         END IF
''         IF NOT ISNULL(rFa!KtrAbrB) THEN  ' bei Privaten
''          op = Format$(3 + 4 + Len(rFa!KtrAbrB), "000") + "4106" + rFa!KtrAbrB
''          Print #327, ZSU(op)
''         END IF
''         IF NOT ISNULL(rFa!AbrAr) THEN  ' bei Privaten
''          op = Format$(3 + 4 + Len(rFa!AbrAr), "000") + "4107" + rFa!AbrAr
''          Print #327, ZSU(op)
''         END IF
''         IF rFa!SchGr <> "90" THEN
''          op = Format$(3 + 4 + 8, "000") + "4109" + Format$(rFa!lVorl, "ddmmyyyy")
''          Print #327, ZSU(op)
''          op = Format$(3 + 4 + 4, "000") + "4110" + Format$(rFa!lVorl, "hhmm")
''          Print #327, ZSU(op)
''          op = Format$(3 + 4 + Len(rFa!IK), "000") + "4111" + rFa!IK
''          Print #327, ZSU(op)
''          IF NOT ISNULL(rFa!KVKs) THEN ' bei Pat_id 43
''           op = Format$(3 + 4 + Len(rFa!KVKs), "000") + "4112" + rFa!KVKs
''           Print #327, ZSU(op)
''          END IF
''          IF NOT ISNULL(rFa!KVKserg) THEN
''           op = Format$(3 + 4 + Len(rFa!KVKserg), "000") + "4113" + rFa!KVKserg
''           Print #327, ZSU(op)
''          END IF
''         END IF
'''         FGeb = format$(rFa!GebOr, "ddmmyyyy")
'''         IF mid$(FGeb, 5, 2) = "20" THEN
'''          IF n!GebDat > Date THEN FGeb = LEFT(FGeb, 4) & "19" & right$(FGeb, 2)
'''         END IF
'''         IF NOT ISNULL(FGeb) THEN ' bei Privaten
''         IF NOT ISNULL(rFa!GebOr) THEN
'''          op = format$(3 + 4 + Len(FGeb), "000") + "4121" & FGeb
''          op = Format$(3 + 4 + Len(rFa!GebOr), "000") + "4121" & rFa!FGeb
''          Print #327, ZSU(op)
''         END IF
''         IF NOT ISNULL(rFa!AbrGb) THEN ' bei Privaten
''          op = Format$(3 + 4 + Len(rFa!AbrGb), "000") + "4122" + rFa!AbrGb
''          Print #327, ZSU(op)
''         END IF
''         op = Format$(3 + 4 + Len("TM#" + IIF(ISNULL(rFa!TMFNr), vns, rFa!TMFNr)), "000") + "4144" + "TM#" + IIF(ISNULL(rFa!TMFNr), vns, rFa!TMFNr)
''         Print #327, ZSU(op)
''         op = Format$(3 + 4 + 8, "000") + "4150" + IIF(rFa!BhFB = 0, "00000000", Format$(rFa!BhFB, "ddmmyyyy"))
''         Print #327, ZSU(op)
''         op = Format$(3 + 4 + 8, "000") + "4151" + IIF(rFa!BhFE1 = 0, "00000000", Format$(rFa!BhFE1, "ddmmyyyy"))
''         Print #327, ZSU(op)
''         op = Format$(3 + 4 + 8, "000") + "4152" + IIF(rFa!BhFE2 = 0, "00000000", Format$(rFa!BhFE2, "ddmmyyyy"))
''         Print #327, ZSU(op)
''         IF 1 = 1 THEN
''          Dim Üw$
''          Üw = IIF(ISNULL(rFa!ÜbwV), vns, rFa!ÜbwV)
''          IF Üw <> vns THEN
''           op = Format$(3 + 4 + Len(Üw), "000") + "4218" + Üw
''           Print #327, ZSU(op)
''          END IF
''          Üw = IIF(ISNULL(rFa!AndÜw), vns, rFa!AndÜw)
''          IF Üw <> vns THEN
''           op = Format$(3 + 4 + Len(Üw), "000") + "4219" + Üw
''           Print #327, ZSU(op)
''          END IF
''          IF NOT ISNULL(rFa!ÜWZiel) AND rFa!ÜWZiel <> vns THEN
''           op = Format$(3 + 4 + Len(rFa!ÜWZiel), "000") + "4220" + rFa!ÜWZiel
''           Print #327, ZSU(op)
''          END IF
''         END IF ' 1 = 1
''         op = Format$(3 + 4 + 2, "000") + "4239" + rFa!SchGr
''         Print #327, ZSU(op)
''         IF NOT ISNULL(rFa!PGeb) THEN
''          FGeb = Format$(CStr(rFa!PGeb), "ddmmyyyy")
''          IF Mid$(FGeb, 5, 2) = "20" THEN
''           IF n!GebDat > Date THEN FGeb = LEFT(FGeb, 4) & "19" & Right$(FGeb, 2)
''          END IF
''          Üw = "TM#" + FGeb
''          op = Format$(3 + 4 + Len(Üw), "000") + "4401" + Üw
''         END IF
''         FGeb = Format$(rFa!PGebErg, "ddmmyyyy")
''         IF Mid$(FGeb, 5, 2) = "20" THEN
''          IF n!GebDat > Date THEN FGeb = LEFT(FGeb, 4) & "19" & Right$(FGeb, 2)
''         END IF
''         IF NOT ISNULL(FGeb) THEN
''          Üw = "TM#" & FGeb
''          op = Format$(3 + 4 + Len(Üw), "000") + "4402" + Üw
''         END IF
''         IF (NOT ISNULL(rFa!GOÄKatNr) AND rFa!GOÄKatNr <> vns) OR (NOT ISNULL(rFa!GOÄKatName) AND rFa!GOÄKatName <> vns) THEN
''          Üw = "TM#" + CStr(rFa!GOÄKatNr) + "#" + CStr(rFa!GOÄKatName)
''          op = Format$(3 + 4 + Len(Üw), "000") + "4580" + Üw
''          Print #327, ZSU(op)
''         END IF
''' IF 1 = 0
''         Print #327, "01380006200"
''         Print #327, "014810000039"
''         op = Format$(3 + 4 + Len(CStr(q!Pat_id)), "000") + "3000" + CStr(q!Pat_id)
''         Print #327, ZSU(op)
''         Print #327, "01380006100"
''         Print #327, "014810000369"
''         op = Format$(3 + 4 + Len(CStr(q!Pat_id)), "000") + "3000" + CStr(q!Pat_id)
''         Print #327, ZSU(op)
''         IF 1 = 0 THEN 'Auswirkung bisher nicht geprüft 31.7.05 (3x)
''          op = Format$(3 + 4 + Len(n!NVorsatz), "000") + "3100" + CStr(n!NVorsatz)
''          Print #327, ZSU(op)
''         END IF
''         op = Format$(3 + 4 + Len(n!Nachname), "000") + "3101" + CStr(n!Nachname)
''         Print #327, ZSU(op)
''         op = Format$(3 + 4 + Len(n!Vorname), "000") + "3102" + CStr(n!Vorname)
''         Print #327, ZSU(op)
''         FGeb = Format$(n!GebDat, "ddmmyyyy")
''         IF Mid$(FGeb, 5, 2) = "20" THEN
''          IF n!GebDat > Date THEN FGeb = LEFT(FGeb, 4) & "19" & Right$(FGeb, 2)
''         END IF
''         op = Format$(3 + 4 + 8, "000") + "3103" + FGeb
''         Print #327, ZSU(op)
''         IF NOT ISNULL(n!Versichertennummer) THEN
''          op = Format$(3 + 4 + Len(n!Versichertennummer), "000") + "3105" + n!Versichertennummer
''          Print #327, ZSU(op)
''         END IF
'''     op = format$(3 + 4 + Len(n!Straße), "000") + "3107" + n!Straße
'''     Print #327, ZSU(op)
'''     op = format$(3 + 4 + Len(n!Plz), "000") + "3112" + n!Plz
'''     Print #327, ZSU(op)
'''     op = format$(3 + 4 + Len(n!Ort), "000") + "3113" + n!Ort
'''     Print #327, ZSU(op)
''         IF NOT ISNULL(n!KVKStatus) AND n!KVKStatus > 0 AND Trim$(n!KVKStatus) <> vns THEN
''          op = Format$(3 + 4 + Len(n!KVKStatus), "000") + "3108" + n!KVKStatus
''          Print #327, ZSU(op)
''         END IF
''         op = Format$(3 + 4 + 1, "000") + "3110" + IIF(n!Geschlecht = "m", "1", IIF(n!Geschlecht = "w", "2", vns))
''         Print #327, ZSU(op)
''' 3610
''         op = Format$(3 + 4 + 8, "000") + "3610" + Format$(n!AufnDat, "ddmmyyyy")
''         Print #327, ZSU(op)
'''( 3626
''' 3629
''' 3626
''' 3629)
''' 3630
''         IF 1 = 0 THEN ' führt zu doppelten HausarztEintraegen
''          IF NOT ISNULL(n!KVNr) THEN
''           op = Format$(3 + 4 + Len(n!KVNr), "000") + "3630" + CStr(n!KVNr)
''           Print #327, ZSU(op)
''          END IF
''         END IF ' 1 = 0
''' 3631
''         op = Format$(3 + 4 + Len(Trim$(CStr(IIF(ISNULL(n!Weggeldzone), vns, n!Weggeldzone)))), "000") + "3631" + Trim$(CStr(IIF(ISNULL(n!Weggeldzone), vns, n!Weggeldzone)))
''         Print #327, ZSU(op)
''        END IF ' 1 = 0
'       END IF ' not n.bof
'#END IF
'' jetzt kommen mit 5999 beginnend die Einzeldiagnosen
'       altpat_id = q!pat_id
'      END IF ' not rFa.BOF THEN ' Not f.NoMatch THEN
'     END IF ' q!Pat_id <> altpat_id THEN
'     BDT.DAdd "3649", aktdat
'     BDT.TAdd "6201", aktdat
'     BDT.TAdd "6203", "TM#?##"
'     Dim ICD$, DiagSi$, DiagText$
'     ICD = q!ICD
'     DiagSi = "G"
'     DiagText = q!Diagnose
'     IF NOT ISNULL(q!ICD) THEN
'      IF q!ICD <> vNS THEN
'       IF InStrB("VGZA", Right$(q!ICD, 1)) > 0 AND Right$(q!ICD, 1) <> vNS THEN
'        DiagSi = Right$(q!ICD, 1)
'        ICD = LEFT(q!ICD, Len(q!ICD) - 1)
'        SELECT CASE DiagSi
'         Case "V": DiagText = LTrim$(replace$(DiagText, "V.a.", vNS))
'         Case "Z": DiagText = LTrim$(replace$(DiagText, "Z.n.", vNS))
'         Case "A": DiagText = LTrim$(replace$(DiagText, "Ausschluss ", vNS))
'        END SELECT
'       END IF
'      END IF
'     END IF
'     Dim BDTFeld$
'     IF ISNULL(q!nurquart) THEN
'      BDTFeld = "3650"
'     ElseIf q!nurquart = 0 THEN
'      BDTFeld = "3650"
'     Else
'      BDTFeld = "6000"
'     END IF
'     BDT.SAdd BDTFeld, DiagText
'     BDT.SAdd "6001", ICD
'     BDT.SAdd "6003", DiagSi
'     BDT.SAdd "6010", "TM#Wahr"
'     BDT.SAdd "6011", "TM#ja"
'     DZahl = DZahl + 1
'' jetzt kommt der Diagnoseneintrag in die Tabelle
''     rDT.Index = "DiagSuch" ' pat_id, icd, diagsicherheit, diagseite
''     rDT.Seek "=", n!Pat_id, Q!ICD
''#If False THEN
'     SET rDT = Nothing
''     rDT.Open "SELECT * FROM `diagnosen` WHERE pat_id = " & q!Pat_id & " AND icd = " & q!ICD, DBCn, adOpenDynamic, adLockOptimistic
''     IF Not rDT.BOF THEN
'     Dim rAF&, rAfL&
'     IF Not obTest THEN
'      InsKorr DBCn, "INSERT INTO `diagnosen`(pat_id,gesName,ICD,diagdatum,diagsicherheit,diagtext,obdauer,aktzeit) VALUES(" & q!Pat_id & ",'" & UmwfSQL(q!Name) & "','" & q!ICD & "'," & DatFor_k(aktdat) & ",'" & DiagSi & "','" & DiagText & "'," & IIF(BDTFeld = "3650", 1, 0) & "," & DatFor_k(BDT.üzpt) & ")", rAF
'      IF rAF <> 1 THEN
'       MsgBox "Fehler beim Diagnoseneeinfügen für Pat. " & q!Pat_id & " (" & q!GesName & ")" & vbCrLf & "ICD: " & q!ICD & vbCrLf & "Diagtext:" & q!DiagText & vbCrLf & "Datum: " & DatFor_k(aktdat) & rAF & " Datensätze eingefügt"
'      END IF
'      IF LenB(q!ICD) <> 0 AND LenB(q!Diagnose) <> 0 THEN
'       Call myEFrag("UPDATE `diagnosenexport` SET status = '" & übertragen & "' WHERE id = " & q!ID, rAF)
'       IF rAF <> 1 THEN
'        MsgBox "Fehler beim Statussetzen in `diagnosenexport` für ID: " & q!ID & rAF & " Datensätze gesetzt"
'       END IF
'       InsKorr DBCn, "INSERT INTO `diagnosen exportiert`(pat_id,datum,icd,diagnose,übertragen) VALUES(" & q!Pat_id & "," & DatFor_k(aktdat) & ",'" & q!ICD & "','" & DiagText & "'," & DatFor_k(BDT.üzpt) & ")", rAF
'       IF rAF > 0 THEN
'        Call myEFrag("DELETE FROM `fuerdiagexp` WHERE id = " & q!ID, rAfL)
'        IF rAfL <> 1 THEN
'         MsgBox "Fehler beim Löschen aus `fuerdiagexp` von " & q!Pat_id & " (" & q!GesName & ")" & vbCrLf & "ICD: " & q!ICD & vbCrLf & "Diagtext:" & q!DiagText & vbCrLf & "Datum: " & DatFor_k(aktdat) & rAfL & " Datensätze gelöscht"
'        END IF
'       END IF
'       IF rAF <> 1 THEN
'        MsgBox "Fehler beim Eintragen in `diagnosen exportiert` von " & q!Pat_id & " (" & q!GesName & ")" & vbCrLf & "ICD: " & q!ICD & vbCrLf & "Diagtext:" & q!DiagText & vbCrLf & "Datum: " & DatFor_k(aktdat) & rAF & " Datensätze eingetragen"
'       END IF
'      END IF ' q!icd <> vns AND q!Diagnose <> vns THEN
'      Call dynDiag(q!Pat_id) ' 12.7.08
'     END IF ' obTEst
''END IF
''#END IF
''   END IF ' q!pat_id = 77
'   q.MoveNext
'  Loop ' do while not q.eof
''  Dim Bm$
''  ON Error Resume Next
''  Bm = Forms("Anamnesebogen").Recordset.Bookmark
'  ON Error GoTo fehler
''  SET Q = Dtb.OpenRecordset("SELECT distinct(pat_id) FROM `" + QMdbAkt + "`.`diagnosenexport`")
'  SET q = Nothing
''  myFrag q, "SELECT distinct(pat_id) FROM `diagnosenexport`"
''  Do While Not q.EOF
''   Call dynDiag(q!Pat_id)
''   q.MoveNext
'''   ON Error Resume Next
'''   Forms("Anamnesebogen").Recordset.FindFirst ("pat_id = " & CStr(q!Pat_id))
'''   Forms("Anamnesebogen").Requery
'''   ON Error GoTo fehler
''  Loop ' do while not q.eof
''  ON Error Resume Next
'''  Forms("Anamnesebogen").Repaint
'''  Forms("Anamnesebogen").Recordset.Bookmark = Bm
''  ON Error GoTo fehler
' END IF ' Not q.BOF THEN
'' Close #327
 Call BDT.Schreib
 MsgBox dzahl & " Diagnosen übertragen in die Datei '" & BDT.DMPImp & "'"
'' DoCmd.Close acTable, "X00D"
'' IF FSO Is Nothing THEN SET FSO = CreateObject("scripting.FileSystemObject")
 On Error Resume Next
 FSO.DeleteFile REPLACE$(BDT.DMPImp, ".BDT", ".txt"), True
 On Error GoTo fehler
 FSO.CopyFile BDT.DMPImp, REPLACE$(BDT.DMPImp, ".BDT", ".txt"), True
 'Shell "cmd /c copy /y " + """" + z + """" & " " & """" + replace$(z, ".BDT", ".txt") + """"
 'Shell "notepad " & vns & z + vns
'' DoCmd.OpenQuery "Leistexp 2 Abfrage"
'' DoCmd.OpenQuery "LeistExp übertragen Abfrage"
'' Call syscmd(4, "Verknüpfe Exportat")
'' Dim con$, td AS DAO.TableDef, Fil AS File
'' SET Fil = FSO.GetFile(replace$(Z, ".BDT", ".txt"))
'' con = "Text;DSN=Labor;FMT=Fixed;HDR=NO;IMEX=2;CharacterSet=1252;DATABASE=" + hVerz + ";TABLE=" + replace$(Fil.ShortName, ".TXT", vns) + "#txt"
'' SET td = Dtb.CreateTableDef("DiagExpVerknue", dbAttachExclusive, Fil.ShortName, con)
'' ON Error Resume Next
'' DoCmd.Close acTable, "DiagExpVerknue", acSaveNo
'' Dtb.Execute "DROP TABLE `" + "DiagExpVerknue" + "`;"
'' ON Error GoTo fehler
'' Dtb.TableDefs.Append td
'' Call BDTVerschönere
'' Call syscmd(acSysCmdSetStatus, "Fertig mit Diagnosen exportieren")
' GoTo schluss
'eintragsfehler:
' Dim tonRunde%
' For tonRunde = 1 To 10
'  Call Sound(WinDir + "\media\Windows XP-Standard.wav")
' Next tonRunde
' Call VerzPrüf(üVerz)
' Open üVerz + "Eintragsfehler" For Append AS #26
' Print #26, CStr(rDT!pat_id) & " " & CStr(rDT!GesName) & " " & IIF(ISNULL(rDT!ICD), vNS, rDT!ICD) & " " & IIF(ISNULL(rDT!DiagText), vNS, rDT!DiagText)
' Close #26
' Resume Next
'schluss:
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DiagnosenExport/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DiagnosenExport()

#If zutesten Then
' aufgerufen nirgends
Function alldynDiag()
 Lese.ProgStart
 Dim rsAnam As New ADODB.Recordset
 myFrag rsAnam, "SELECT pat_id FROM `anamnesebogen` WHERE pat_id = 52996"
 Do While Not rsAnam.EOF
  Call dynDiag(CStr(rsAnam!Pat_ID))
'  Debug.Print "Fertig mit " & rsAnam!Pat_id
  rsAnam.MoveNext
 Loop
 MsgBox "Fertig mit alldynDiag!"
End Function ' alldynDiag() und doDiagnosenExport
#End If

' aufgerufen in alldynDiag und doDiagnosenExport
Function dynDiag$(Pat_ID$) ' für DiagnosenExport
 On Error GoTo fehler
 Dim DBDiagString$, DiagTab() As CString
 Dim rAf&
 DBDiagString = DiagString$(Pat_ID, DiagTab)
 If InStrB(DBDiagString, "'") <> 0 Then
  DBDiagString = REPLACE$(DBDiagString, "'", "''")
 End If
 Call myEFrag("UPDATE `anamnesebogen` SET diagnosen = '" & DBDiagString & "' WHERE pat_id = " & Pat_ID, rAf)
' SET rsAnam = TabÖff("Anamnesebogen", "pat_id")
' IF Not rsAnam.BOF THEN
'  rsAnam.Seek "=", Pat_id
'  IF Not rsAnam.NoMatch THEN
'   rsAnam.Edit
'   rsAnam!Diagnosen = Diagstring$(Pat_id)
'   IF rsAnam!Pat_id = 0 THEN
'    rsAnam.CancelUpdate ' kommt auf mir noch unbekanntem Weg zustande
'   Else
'    rsAnam.Update
'   END IF
'  END IF
' END IF
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dyndiag/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dyndiag$

Function lFDat(Pat_ID&, Optional nurKasse%) As Date ' letztes Falldatum, falls später als heute, dann heute
' dtbInit
' lFDat = Dtb.OpenRecordset("SELECT MAX(bhfe1) AS bhfe FROM `faelle` WHERE pat_id = " + CStr(Pat_id))!BhFE
 Dim rs As New ADODB.Recordset
 myFrag rs, "SELECT MAX(bhfe1) AS bhfe FROM `faelle` WHERE pat_id = " & Pat_ID & IIf(nurKasse <> 0, " AND schgr<>90", vNS)
 If Not rs.BOF Then
  lFDat = rs!BhFE
 End If
 If lFDat > Now Then lFDat = Now
End Function ' lFDat(pat_id&) As Date

Function test_fdübertrag()
 Dim q As New ADODB.Recordset, z As New ADODB.Connection, i%, rAf&
 Dim JCn$
 JCn$ = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & StACCDB & "';"
 z.Open JCn
 Call Lese.ProgStart
 myFrag q, "SELECT * FROM `fuerdiagexp`"
 Do While Not q.EOF
  InsKorr z, "INSERT INTO `fuerdiagexp`(name,pat_id,icd,diagnose,nurquart,zeitpunkt) VALUES('" & UmwfSQL(q!name) & "'," & q!Pat_ID & ",'" & q!ICD & "','" & q!Diagnose & "'," & q!nurquart & "," & DatFor_k(Now()) & ")", rAf
  q.Move 1
 Loop
End Function ' test_fdübertrag()

