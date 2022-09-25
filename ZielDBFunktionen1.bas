Attribute VB_Name = "ZielDBFunktionen"
Option Explicit
Declare Function Tüt& Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long)
Public Const Verspätung% = 20 ' wie lange ins neue Quartal die alte Abrechnung geht
Const obAdo% = 0
' Inhalte folgender Variablen werden von der Procedure MedPlanNr an die sie aufrufenden weitergegeben:
Public Med$(), Dos$(), MedNr%, MedZahl%

Public lBehDat As Date
Public Const artspez$ = "(art in (""notiz"",""ni"",""telef"",""gs"",""rz"",""ga"",""ag"",""hz"",""ts"",""cr"",""ep"",""tst"",""wr"",""ek"",""ph"",""bga"",""gstel"",""rz"",""ga"",""ag"",""hz"",""ts"",""ke"")"
Public sql$, sql1$ ' SQL-Text für alle möglichen Abfragen
Dim QMdbAkt$, nzw$
Dim DmPStr$
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
 Nie%
 Früher%
 Aktuell%
 VorLangem%
End Enum
Public Enum TherapieArt
 offen
 diät
 oad
 komb
 ct
 ict
 csii
End Enum
'Dim rDT As DAO.Recordset
'Dim rFot As DAO.Recordset ' solange Fotos noch in Access-Datenbank
Dim sstr$
Public Const Titel$ = "KV-Ärzte raussuchen lassen"
Public WinDir$
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096
Declare Function OpenClipboard& Lib "user32" (ByVal hWnd&)
Declare Function CloseClipboard& Lib "user32" ()
Declare Function GetClipboardData& Lib "user32" (ByVal wFormat&)
Declare Function GlobalLock& Lib "kernel32" (ByVal hMem&)
Declare Function GlobalUnlock& Lib "kernel32" (ByVal hMem&)
Declare Function lstrcpy& Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any)
Declare Function GetWindowsDirectory& Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer$, ByVal nSize&)
Public lDMPPat_id&
Public Const forminhalt$ = "SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, form_abk FROM (((forminhfeld LEFT JOIN forminhkopf ON forminhfeld.foid=forminhkopf.foid) LEFT JOIN formulare ON formulare.formid=forminhkopf.form_id) LEFT JOIN forminhaltfeld ON forminhfeld.feldvw=forminhaltfeld.feldvw) LEFT JOIN forminhaltfeldinh ON forminhfeld.feldinhvw=forminhaltfeldinh.feldinhvw "
'Public LeistBDT$
Public Enum DMPStat
 unauff
 ndok
 auff
End Enum
Public Enum DMPSeite
 unbek
 gleich
 re
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
Public Type dmptyp
 Pat_id As Long
 NachName As String
 Vorname As String
 dtyp As String ' Diabetestyp, "1", "2"
 daseit As String
 dspsy As Boolean ' diabetesspezifische Symptome
 dspmed As Boolean ' diagnosespezifische Medikation
 FE(17) As Integer ' Folgeerkrankungen bekannt
 FEn(17) As Integer ' Folgeerkrankungen
 SE(2) As Date ' Schulungen empfohlen
 kSE As Date ' bei letzter Doku keine Schulung empfohlen
 SW(5) As Date ' Schulungen wahrgenommen
 ds(2) As Date ' DiabSchulung
 HS(2) As Date ' HyperSchulung
 VorM(12) As Integer
 obGlib As AntidiabMedType
 obmetf As AntidiabMedType
 obGlucI As AntidiabMedType
 obSHGlin As AntidiabMedType
 obGlit As AntidiabMedType
 obSonstAD As AntidiabMedType
 insz As Integer
 obIns As Boolean
 obAnal As Boolean
 obHMG As Boolean
 obAntihyp As Boolean
 obACEH As Boolean
 obBetabl As Boolean
 obThro As Boolean
 Tabak As Boolean
 kgr As Integer ' Körpergröße
 gewi As Single ' Gewicht
 bmi As Single ' Body Mass Index
 puls As DMPStat ' Pulsstatus
 sens As DMPStat ' Sensibiliätsprüfung
 PrRR As String ' Blutdruck
 RRsyst As Integer
 RRdiast As Integer
 bekHb As Single ' HbA1c
 Crea As Single ' Kreatinin
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
 krZAn As Integer ' Zahl der stationären Aufenthalte wegen Diabetes aus Anamnese
 krZKK As Integer ' Zahl der stationären Aufenthalte wegen Diabetes aus Karteikarte
 obVorb As Boolean ' ob Vorbefund da
 obNI As Boolean ' ob Niereninsuffizienz
 neuDial As Boolean ' ob neu Dialyse
 neuErbl As Boolean ' ob neu Erblindung
 neuAmp As Boolean ' ob neu Amputation
 neuMI As Boolean ' ob neu Myokardinfarkt
 ernb As Boolean ' Ernährungsberatung
 hbEmpf As DMPEmpf ' HbA1c-Empfehlung
 rrEmpf As DMPEmpf ' Blutdruckempfehlung
 aug As DMPaug ' Augenuntersuchung
 tart As TherapieArt ' Therapieart
 obSchw As Boolean
End Type
Function DMPString$(pid&, üdt As dmptyp, Optional ohneVorDMP%)
 Const ErgebDatei$ = aVerz & "DMP.txt"
' Dim dt.obglib As AntidiabMedType, dt.obmetf As Boolean, dt.obgluci As Boolean, dt.obshglin As Boolean, dt.obglit As Boolean, dt.obsonstAD as boolean, _
     dt.obins As Boolean, dt.obanal As Boolean, dt.obhmg As Boolean, dt.obantihyp As Boolean, dt.obthro As Boolean
'Dim rEi As DAO.Recordset, rLaU As DAO.Recordset, rKH As DAO.Recordset
Dim raEi As New ADODB.Recordset, raLau As New ADODB.Recordset, raKH As New ADODB.Recordset
Dim lapp As New ADODB.Recordset
'Dim ralauSql$
Dim DT As dmptyp
'Dim rdo As DAO.Recordset
'Dim rado As New ADODB.Recordset
'Dim rAna As DAO.Recordset
Dim raAna As New ADODB.Recordset
Dim raDT As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rAna As New ADODB.Recordset
Dim GesNa$
Dim DMSchulz%, HSchulz%, SchulStr$
Dim Tbl
Dim lsql$
'Dim RRsyst%, RRdiast%
Dim fiabfr$
'Dim obVorb% ' ob vorbefund da
' Beim ersten Aufruf von diI muß Pid enthalten sein
If diI("Z34", pid) Or diI("Z33", pid) Then DT.obSchw = True
If diI("E10") Then DT.dtyp = "1" Else If diI("E11") Then DT.dtyp = "2"
If Not ohneVorDMP Then
fiabfr = "SELECT Pat_ID, FID, form_abk,Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, form_abk FROM (((forminhfeld LEFT JOIN forminhkopf ON forminhfeld.foid=forminhkopf.foid) LEFT JOIN formulare ON formulare.formid=forminhkopf.form_id) LEFT JOIN forminhaltfeld ON forminhfeld.feldvw=forminhaltfeld.feldvw) LEFT JOIN forminhaltfeldinh ON forminhfeld.feldinhvw=forminhaltfeldinh.feldinhvw where form_abk like ""dmpdtyp%"" and feldinh = ""X"" and pat_id = " & pid & " order by zeitpunkt"
raAna.Open fiabfr, dbcn, adOpenDynamic, adLockReadOnly
Do While Not raAna.EOF
 DT.obVorb = True
 Select Case raAna!Feld
  Case "Folgeerkrankung"
   DT.FE(raAna!FeldNr) = True
  Case "SchulEmpfohlen" ' 0 = Diabetes, 1 = Hypertonie, 2 = keine
   DT.SE(raAna!FeldNr) = raAna!Zeitpunkt
  Case "SchulWahrgenommen" ' Schulungen bereits vor Einschreibung wahrgenommen: 0 = Diabetes, 1 = Hypertonie, 2 = keine
   DT.SW(raAna!FeldNr) = raAna!Zeitpunkt
  Case "DiabSchulung" ' empfohlen Schulungen wahrgenommen, Diabetes: 0 = ja, 1 = nein, 2 = war aktuell nicht möglich
   DT.ds(raAna!FeldNr) = raAna!Zeitpunkt
  Case "HyperSchulung" ' empfohlen Schulungen wahrgenommen, Hypertonie: 0 = ja, 1 = nein, 2 = war aktuell nicht möglich
   DT.HS(raAna!FeldNr) = raAna!Zeitpunkt
  Case "LetzteSchulung" ' bei letzter Doku war keine Schulung empfohlen
   DT.kSE = raAna!Zeitpunkt
  Case "Glibenclamid"
   DT.VorM(0) = raAna!FeldNr
  Case "Metformin"
   DT.VorM(1) = raAna!FeldNr
  Case "Glucosidase"
   DT.VorM(2) = raAna!FeldNr
  Case "Sulfonylharnstoffe"
   DT.VorM(3) = raAna!FeldNr
  Case "Glitazone"
   DT.VorM(4) = raAna!FeldNr
  Case "Insulin"
   DT.VorM(5) = raAna!FeldNr
  Case "Insulinanaloga"
   DT.VorM(6) = raAna!FeldNr
  Case "HMG"
   DT.VorM(7) = 0
  Case "Antihypertensive", "Antihypertensiva"
   DT.VorM(8) = 0
  Case "Thrombozyten"
   DT.VorM(9) = 0
  Case "ACEHemmerDM2"
   DT.VorM(10) = 0
  Case "BetablockerDM2"
   DT.VorM(11) = 0
  Case "AntihypertensivDM2"
   DT.VorM(12) = 0
 End Select
 raAna.Move 1
Loop
raAna.Close
End If 'not ohneVorDMP
DT.RRsyst = 0
DT.RRdiast = 0
On Error GoTo fehler
'Set rMA = TabÖff("MedArten", "Medikament")
If obAdo Then
 raAna.Open "select * from anamnesebogen where pat_id = " & pid, dbcn, adOpenKeyset, adLockReadOnly
 rAna.Open "select * from namen where pat_id = " & pid, dbcn, adOpenKeyset, adLockReadOnly
Else
' Set rEi = TabÖff("eintraege", "Auswahl")
''Set rlau = TabÖff("Labor", "WertSuch")
' Set rDT = TabÖff("Diagnosen")
' Set rKH = TabÖff("KHEinweis", "Auswahl")
' Set rAna = TabÖff("Anamnesebogen", "pat_id")
' Call raEi.Open("select * from eintraege where pat_id = " & PID, DBCn, adOpenDynamic, adLockReadOnly)
 'Call raDT.Open("select * from diagnosen where pat_id = " & PID, DBCn, adOpenDynamic, adLockReadOnly)
 'Call raKH.Open("select * from kheinweis where pat_id = " & PID, DBCn, adOpenDynamic, adLockReadOnly)
' Call raAna.Open("select -dialyse as j_dialyse, a.* from anamnesebogen a where pat_id = " & pid, DBCn, adOpenDynamic, adLockReadOnly)
 Call raAna.Open("select * from anamnesebogen where pat_id = " & pid, dbcn, adOpenDynamic, adLockReadOnly)
 Call rAna.Open("select * from namen where pat_id = " & pid, dbcn, adOpenDynamic, adLockReadOnly)
' DoCmd.Save acForm, Anmnb geht leider nicht
'If PID = 0 Then
' Dim tonRunde%
' For tonRunde = 1 To 10
'  Call Sound(WinDir + "\media\Windows XP-Standard.wav")
' Next tonRunde
' MsgBox "Achtung: DMP-Infos mit PID=0!"
' On Error Resume Next
' PID = Forms!Anamnesebogen!Pat_id
' On Error GoTo fehler
' If PID = 0 Then
'  For Each tbl In Dtb.TableDefs
'   DoCmd.Close acTable, tbl.Name, acSaveYes
'  Next
'  DoCmd.OpenForm Dtb.Containers(2).Documents(0).Name
'  DoCmd.Maximize
'  PID = Forms!Anamnesebogen!Pat_id
' End If
'End If
End If
DT.Pat_id = pid
If IsNull(raAna!NachName) Then ' so am 2.9.08 ereignet
 DT.NachName = rAna!NachName
 DT.Vorname = rAna!Vorname
Else
 DT.NachName = raAna!NachName
 DT.Vorname = raAna!Vorname
End If
Pat_id = pid
'Dim sqllabor$
'sqllabor$ = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, Abkü, Langtext, Wert, Einheit, Kommentar, AbsPos, AktZeit FROM laborlangtext INNER JOIN (laborkommentar INNER JOIN laborneu ON laborkommentar.KommentarVW = laborneu.KommentarVW) ON laborlangtext.LangtextVW = laborneu.LangtextVW where pat_id = " & CStr(pid)
''Const sqllabor$ = "SELECT Pat_ID, ZeitPunkt, FertigStGrad, Abkü, Langtext, Wert, Einheit, Kommentar, AbsPos, AktZeit FROM laborlangtext INNER JOIN (laborneu) ON laborlangtext.LangtextVW = laborneu.LangtextVW"
'lsql = "select Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" as NB from (" & sqllabor & ") as labor where wert <> """" and not isnull(wert)"
'lsql = "select Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" as NB from (" & sqllabor & ") as labor where wert <> """" and not isnull(wert) UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar,Normbereich as NB " + _
 "FROM laborxus LEFT JOIN laborxwert ON laborxus.RefNr=laborxwert.RefNr " + _
 "WHERE pat_id = " & DT.Pat_id & " and wert <> """" and not isnull(wert) and not exists (select * from laborneu where pat_id = " & DT.Pat_id & " and abkü = laborxwert.Abkü and wert = laborxwert.wert and zeitpunkt > laborxus.Eingang-3 and zeitpunkt < laborxus.Eingang+6)"
lsql = "select * from labor1 where pat_id = " & Pat_id & " union select * from labor2 where pat_id = " & Pat_id
'ralauSql = "select * from (" & lsql & ") as innen order by zeitpunkt desc"

'Call raLaU.Open(ralauSql, DBCn, adOpenDynamic, adLockReadOnly)
'Do While Not raLaU.EOF
' Debug.Print raLaU!Abkü & " " & raLaU.Fields(0) & " " & raLaU.Fields(1) & " " & raLaU.Fields(2) & " " & raLaU.Fields(3) & " " & raLaU.Fields(4) & " " & raLaU.Fields(5)
' raLaU.Move 1
'Loop

'Set ralau = Dtb.OpenRecordset("select * from [" + QMdbAkt + "].laborunion where pat_id = " + CStr(dt.pat_id) + " order by zeitpunkt desc")
'With Forms!anamnesebogen
GesNa = GesNam(rAna) & ", *" & Format$(raAna!gebdat, "dd.mm.yyyy") & " (Pat'nr: " & rAna!Pat_id & ")"

With raAna
 DmPStr = "DMP-Informationen vom " & Format$(Date, "DD.MM.YY") & " zu " & vbTab & vbTab & vbTab & vbTab & vbCrLf & GesNa + ":"
 
 If Not IsNull(raAna!Vorgestellt) Then
  DT.daseit = zuJahr(daseit(raAna), raAna!Vorgestellt)
  TabPr "D.m. Typ " + dtyp(raAna!Diabetestyp) & IIf(DT.dtyp <> raAna!Diabetestyp, " (laut Anamnesebogen; " & DT.dtyp & " laut ICD)", "") & " seit:", DT.daseit
 End If
 If obPosi(![Unterzucker pM]) Or obPosi(![keto pa]) Or obPosi(![Fremde Hilfe pa]) Or obPosi(![Bewußtlos pa]) Then
  DT.dspsy = True
  TabPr "Diagnosespez. Symptome: ", "ja"
 End If
 If diI("Z34") Or diI("Z33") Then
  TabPr "Schwangerschaft laut ICD-Diagnosen:", "ja"
  DT.obSchw = True
 End If
 Dim vordat As Date '(nur Attrappe)
 Call TherAuskunft(pid, 0, DT.insz, vordat, DT.obIns, DT.obAnal, DT.obGlib, DT.obmetf, DT.obGlucI, DT.obSHGlin, DT.obGlit, DT.obSonstAD, DT.obHMG, DT.obAntihyp, DT.obACEH, DT.obBetabl, DT.obThro)
 'TabPr "Diagnosespez. Medik.: ", IIf(![DiabetesMedikament 1] <> "" And Not IsNull(![DiabetesMedikament 1]), "ja", "nein")
 If DT.dtyp = "1" And Not DT.obIns And Not DT.obAnal Then ' Medikamentenplan fehlt oder keine Dosierung angegeben
  Dim rez As New ADODB.Recordset, rezSQL$
  rezSQL = "select * from (SELECT *,left(medikament,instr(medikament,"" "")-1) as medanf FROM rezepteintraege) as rez left join medarten on rez.medanf = medarten.medikament collate latin1_general_ci where rez.pat_id = 150 and (ins or anal) order by zeitpunkt desc;"
  If Not lies.obMySQL Then rezSQL = Replace(rezSQL, "collate latin1_general_ci ", "")
  rez.Open rezSQL, dbcn, adOpenStatic, adLockReadOnly
  If Not rez.BOF Then
   DT.obIns = rez!InS
   DT.obAnal = rez!AnAl
  End If
  rez.Close
 End If
 DT.dspmed = (DT.obIns Or DT.obAnal Or DT.obGlib Or DT.obmetf Or DT.obGlucI Or DT.obSHGlin Or DT.obGlit Or DT.obSonstAD)
' dt.dspmed = (dt.VorM(0) Or dt.VorM(1) Or dt.VorM(2) Or dt.VorM(3) Or dt.VorM(4) Or dt.VorM(5) Or dt.VorM(6))
 TabPr "Diagnosespez. Medik.: ", IIf(DT.dspmed, "ja", "nein")
 Dim bbk$
 bbk = ""
 'If instrb(!Diagnosen, "yperton") <> 0 Or instrb(!Diagnosen, "ochdru") <> 0 Then bbk = bbk + " Hypertonus,"
 'If diT("yperton") Or diT("ochdru") Then bbk = bbk + " Hypertonus,"
 If diI("I10") Then
  DT.FEn(1) = True
  bbk = bbk + " Hypertonus" & IIf(DT.FE(1), "( bek)", "") & ","
 End If
 If diT("ettstoffw") Or diT("yperchol") Or diI("E78") Then
  DT.FEn(2) = True
  bbk = bbk + " Fettstoffwechselstörung" & IIf(DT.FE(2), "( bek)", "") & ","
 End If
 If diI("I25") Then
  DT.FEn(3) = True
  bbk = bbk + " Koronare Herzerkrankung" & IIf(DT.FE(3), "( bek)", "") & ","
 End If
 If obPosi(!Herzinfarkt) Or diI("I21") Then
  DT.FEn(4) = True
  bbk = bbk + " Herzinfarkt" & IIf(DT.FE(4), "( bek)", "") & ","
 End If
 bbk = bbk & vbCrLf & vbTab & vbTab & vbTab
 If diT("poplex") Or diT("chlaganf") Or (diT("erebral") And diT("schämie")) Or diI("I63") Or diI("I64") Then
  DT.FEn(5) = True
  bbk = bbk + " Schlaganfall" & IIf(DT.FE(5), "( bek)", "") & ","
 End If
 If diI("I7") Then
  DT.FEn(6) = True
  bbk = bbk + " pAVK" & IIf(DT.FE(6), "( bek)", "") & ","
 End If
 If diI("N08.3") And Not obkNeph Then
  DT.FEn(7) = True
  bbk = bbk + " Nephropathie" & IIf(DT.FE(7), "( bek)", "") & ","
 End If
' If obPosi(!j_Dialyse) Or diI("Z49.1") Then
 If obPosi(!Dialyse <> 0) Or diI("Z49.1") Then
  DT.FEn(8) = True
  bbk = bbk + " Nierenersatztherapie" & IIf(DT.FE(8), "( bek)", "") & ","
 End If
 
 If diT("Retinopath") Or diI("H36") Then
  DT.FEn(9) = True
  bbk = bbk + " diab. Retinopathie" & IIf(DT.FE(9), "( bek)", "") & ","
 End If
 bbk = bbk & vbCrLf & vbTab & vbTab & vbTab
 If InStrB(LCase(!Diagnosen), "blind") <> 0 Or diI("H54") Then
  DT.FEn(10) = True
  bbk = bbk + " Blindheit" & IIf(DT.FE(10), "( bek)", "") & ","
 End If
 If diI("G63.2") Or diI("G59.0") Or diI("G99.0") Or diT("europath") Then
  DT.FEn(11) = True
  bbk = bbk + " diab. Neuropathie" & IIf(DT.FE(11), "( bek)", "") & ","
 End If
 If diI("L89") Or diI("M14.6") Then
  DT.FEn(12) = True
  bbk = bbk + " Diabetisches Fußsyndrom" & IIf(DT.FE(12), "( bek)", "") & ","
 End If
 If obPosi(!Amputation) Or diT("mputat") Then
  DT.FEn(13) = True
  bbk = bbk + " Amputation" & IIf(DT.FE(13), "( bek)", "") & ","
 End If
 If diI("M36.8") Or diI("M14.2") Or InStrB(!Diagnosen, "aculop") <> 0 _
 Or diI("K77.8") Or diI("K71.7") Or diI("K71.6") Or diI("K76.9") Or diI("K76.2") _
 Or diT("atara") > 0 Or diT("arnwegsinf") > 0 Then
  DT.FEn(14) = True
  bbk = bbk + " Sonstige" & IIf(DT.FE(14), "( bek)", "") & ""
 End If
 If diI("J42") Or diI("J44") Or diI("N40") Then ' COPD
  DT.FEn(15) = True
 End If
 If diI("J45") Then ' Asthma
  DT.FEn(16) = True
 End If
 If diI("I50") Then ' Herzinsuffizienz
  DT.FEn(17) = True
 End If
 Dim j%
 DT.FEn(0) = -1
 For j = 1 To UBound(DT.FEn)
  If DT.FEn(j) And Not DT.FE(j) Then
   DT.FEn(0) = 0
   Exit For
  End If
 Next j
 If bbk = "" Then bbk = " keine"
 If Right$(bbk, 1) = "," Then bbk = Left(bbk, Len(bbk) - 1)
 TabPr "Bek. Begleit-/Folgeerk.:", bbk
' Dim TabakExpr$
' If IsNull(!Tabak) Or !Tabak = "" Then
'  TabakExpr = ""
' Else
'  If instrb("jJ", trim$(!Tabak)) > 0 Then
'   TabakExpr = "ja"
'  ElseIf instrb("-nN0", trim$(!Tabak)) > 0 Then
'   TabakExpr = "nein"
'  Else
'   TabakExpr = CStr(!Tabak)
'  End If
' End If
 DT.Tabak = (WieTabak(tfeld(!Tabak), !Pat_id) = Aktuell)
 TabPr "Tabak: ", IIf(DT.Tabak, "ja", "nein")
 If Not IsNull(!Größe) Then DT.kgr = Round(Replace(IIf(!Größe < 10, !Größe * 100, !Größe), ".", ","), 0)
 If Not IsNull(!Größe) Then TabPr "Körpergröße: ", CStr(!Größe) + IIf(!Größe < 10, " m", " cm")
 
 raEi.Open "select * from eintraege where pat_id = " & pid & " and art = ""Gewicht"" order by zeitpunkt desc", dbcn, adOpenDynamic, adLockReadOnly
 If Not raEi.EOF Then
  Dim gewi$
  gewi = Trim$(Replace(Replace(LCase(raEi!Inhalt), "kg", ""), ".", ","))
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
       gewi = Replace(gewi, Chr$(j), "")
      End If
     Next j
    End If
   End If
   gewi = Replace(gewi, ",,", ",")
   If gewi = "" Then
    gewi = "0"
    Exit Do
   End If
   If IsNumeric(gewi) Then Exit Do
  Loop
  If Not IsNumeric(gewi) Then
   DT.gewi = 0
  Else
   DT.gewi = Round(gewi, 0)
  End If
  TabPr "Körpergewicht: ", raEi!Inhalt + IIf(InStrB(raEi!Inhalt, "kg") <> 0, "", " kg") & " (" & CStr(DateValue(raEi!Zeitpunkt)) & ")"
 Else
  If Not IsNull(!Gewicht) Then
   DT.gewi = Round(Replace(!Gewicht, ".", ","), 0) ' ist double
   TabPr "Körpergewicht: ", CStr(!Gewicht) + " kg"
  End If
 End If
    
    If DT.kgr <> 0 Then
     DT.bmi = DT.gewi / (DT.kgr * 1# * DT.kgr)
     If DT.bmi > 0 Then
      Do While DT.bmi < 10
       DT.bmi = DT.bmi * 10
      Loop
      Do While DT.bmi > 100
       DT.bmi = DT.bmi * 0.1
      Loop
     End If
    End If
 
 On Error Resume Next
 ReDim rNa(0)
 On Error GoTo fehler
 rNa(0).Pat_id = pid
 Call usDM(True)
 Dim trp As Boolean, tlp As Boolean, drp As Boolean, dlp As Boolean
 Dim Mfrep%, Mflip%
 Dim KWrep%, KWlip%
 Dim VibrIKrep%, VibrIKlip%
 Dim VibrGZrep%, VibrGZlip%
 Dim summe%
 Dim Nausg$, LUSDat$, obG632%
 Dim MFBegr$, KWBegr$, VibBegr$
 DT.puls = unauff
 If UBound(AbI) = 0 Or (AbI(3) = "" And AbI(2) = "" And AbI(5) = "" And AbI(4) = "") Then
  PulsParse ![Puls Atp], trp, tlp
  PulsParse ![Puls Adp], drp, dlp
  Nausg = " (" & tfeld(![Puls Atp]) & "," & tfeld(![Puls Adp]) & ")"
  LUSDat = !Vorgestellt
  If tfeld(![Puls Atp]) = "" And tfeld(![Puls Adp]) = "" Then DT.puls = ndok
 Else
  PulsParse AbI(3) & "/" & AbI(2), trp, tlp
  PulsParse AbI(5) & "/" & AbI(4), drp, dlp
  Nausg = " (" & AbI(3) & "/" & AbI(2) & ", " & AbI(5) & "/" & AbI(4) & ")"
  LUSDat = Format$(AbIDate, "d.m.yy")
 End If
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
 If (trp And drp) Or (tlp And dlp) Or (trp And dlp) Or (tlp And drp) Then
   DT.puls = auff
 End If
 If DT.puls = auff Then
  TabPr "Pulsstatus: ", "auffällig (" & LUSDat & Nausg
 ElseIf DT.puls = ndok Then
  TabPr "Pulsstatus: ", "nicht dokumentiert (" & LUSDat & Nausg
 Else
  TabPr "Pulsstatus: ", "unauffällig (" & LUSDat & Nausg
 End If
  
 summe = Mfrep + Mflip + KWrep + KWlip + VibrIKrep + VibrIKlip + VibrGZrep + VibrGZlip
 Nausg = "auffällig (" & LUSDat
 obG632 = diI("G63.2")
 If summe > 1 Or obG632 Then
  If Mfrep > 0 Or Mflip > 0 Then Nausg = Nausg & " Monofil.:" & MFBegr & ","
  If KWrep > 0 Or KWlip > 0 Then Nausg = Nausg & " Kalt/W.:" & KWBegr & ","
  If VibrGZrep > 0 Or VibrGZlip > 0 Then Nausg = Nausg & " Vibr." & VibBegr & ","
  If obG632 Then Nausg = Nausg & " Diagn. G63.2"
  DT.sens = auff
  TabPr "Sensibilitätsprüfung: ", Nausg
 ElseIf tfeld(![Monofilamenttest]) = "" And (tfeld(![Kalt-Warm]) = "" Or tfeld(![Kalt-Warm]) = "/5 | /5" Or tfeld(![Kalt-Warm]) = "~/5 | ~/5") And tfeld(![Vibration IK]) = "" And tfeld(![Vibration Großzehe]) = "" Then
  DT.sens = ndok
  TabPr "Sensibilitätsprüfung: ", "nicht dokumentiert"
 Else
  DT.sens = unauff
  TabPr "Sensibilitätsprüfung: ", "un" & Left(Nausg, Len(Nausg) - 2) & ")"
 End If
 DT.PrRR = GetPrRR(raAna, DT.RRsyst, DT.RRdiast)
 If DT.dtyp = "2" Then
  If DT.PrRR <> "" Then
   TabPr "Blutdruck: ", DT.PrRR
  End If
 End If
 DT.bekHb = 0
' ralau.Seek "=", CStr(PID), "HBA1C"
 Dim DMPHbA1c$, DMPCrea$, DMPUAlb$
 Set raLau = Nothing
 raLau.Open "select * from (" & lsql & ") as innen where abkü = ""HBA1C"" order by zeitpunkt desc", dbcn, adOpenDynamic, adLockReadOnly
' raLaU.MoveFirst
' raLaU.Find "Abkü = ""HBA1C"""
 If Not raLau.EOF Then
  If raLau!Wert <> "" Then
   DT.bekHb = Round(CDbl(Replace(raLau!Wert, ".", ",")), 1)
  End If
  DMPHbA1c = raLau!Wert + "% (" + Format$(raLau!Zeitpunkt, "dd/mm/yy") + "), oberer Normwert: 6,2%"
 ElseIf tfeld(![letztes HbA1c]) <> "" Then
  Dim testHbA1c
  testHbA1c = Replace(Replace(![letztes HbA1c], ".", ","), "%", "")
  If IsNumeric(testHbA1c) Then
   DT.bekHb = Round(Replace(testHbA1c, ".", ","), 1)
  End If
  DMPHbA1c = ![letztes HbA1c] + "% " + IIf(tfeld(![gemessen am]) = "", "(auswärtig)", "( " + tfeld(![gemessen am]) + ")")
 End If
' ralau.Seek "=", CStr(PID), "KREA"
 Set raLau = Nothing
 raLau.Open "select * from (" & lsql & ") as innen where abkü in (""KREA"", ""KREA02"", ""KRE02"", ""CREAT"") order by zeitpunkt desc", dbcn, adOpenDynamic, adLockReadOnly
' raLaU.Find "Abkü in (""KREA"",""KREA02"")" 'like ""KREA" & "%" & """

 If Not raLau.EOF Then
  DT.Crea = Round(Replace$(Replace$(raLau!Wert, ".", ","), ";", ","), 1)
  DMPCrea = raLau!Wert + " mg/dl (" + Format$(raLau!Zeitpunkt, "dd/mm/yy") + ")" ', oberer Normwert: 1,3 mg/dl"
 End If
' If DT.dtyp = "1" Then ' 5.10.08: jetzt auch Typ 2
  Set raLau = Nothing
  raLau.Open "select * from (" & lsql & ") as innen where abkü in ('ALBCRE','ALBKRE','ALBQ','ALBU') and (abkü <> 'ALBU' or wert like '%<%') order by zeitpunkt desc", dbcn, adOpenDynamic, adLockReadOnly
' raLaU.Find "Abkü in (""KREA"",""KREA02"")"
  If raLau.EOF Then
   Dim rEintr As New ADODB.Recordset
   rEintr.Open "select * from eintraege where art = ""urin"" and inhalt like ""%micral%"" And Pat_ID = " & DT.Pat_id & " order by zeitpunkt desc", dbcn, adOpenStatic, adLockReadOnly
   If Not rEintr.EOF Then
    DMPUAlb = Mid$(rEintr!Inhalt, InStr(LCase(rEintr!Inhalt), "micral")) & " (" & Format$(rEintr!Zeitpunkt, "dd/mm/yy") & ")"
   End If
  Else
   If InStrB(raLau!Wert, "<") <> 0 Then
    DT.mau = unauff
   Else
    If Round(Replace(Replace(raLau!Wert, "ca.", ""), ".", ","), 1) > 20 Then
     DT.mau = auff
    Else
     DT.mau = unauff
    End If
   End If
   DMPUAlb = raLau!Wert + " mg/dl (" + Format$(raLau!Zeitpunkt, "dd/mm/yy") + "), oberer Normwert: 20 mg/gCrea"
   If IsNumeric(raLau!Wert) Then
    If raLau!Wert > 20 Then
     DT.mau = auff
    Else
     DT.mau = unauff
    End If
   End If
'  End If
 End If
 If DT.dtyp = "2" Then
  TabPr "HbA1c: ", DMPHbA1c
  TabPr "Serum-Kreatinin: ", DMPCrea
  TabPr "Urin-Albumin: ", DMPUAlb ' 5.10.08 Kommentar entfernt
 End If
 'If obPosi(!Beinbefund) Or obPosi(!Ulcera) Or diI("M14") Then
 
 Dim raFa As New ADODB.Recordset
 Dim diIsql$
 diIsql = "select * from diagnosen where pat_id = " & pid & " and (icd like ""M14.6%"" or icd like ""T79.%"" or icd like ""L89.%"" or icd like ""T89.%"" or icd like ""T87.4%"" or icd like ""Z44.1%"") and diagsicherheit in (""G"",""V"",""Z"")" ' M14.6 = Charcot, M14.2 = Arthorpathie
 Dim lddat As Date
 raFa.Open "select max(bhfb) as lddat from faelle where pat_id = " & pid, dbcn, adOpenStatic, adLockReadOnly
 If Not raFa.BOF Then
  lddat = raFa!lddat
  diIsql = diIsql & " and (obdauer or diagdatum >= " & datform(lddat) & ")"
 End If
 Set raDT = Nothing
 Call raDT.Open(diIsql, dbcn, adOpenDynamic, adLockReadOnly)
 Dim obUlc%, FSauf$
 obUlc = obPosi(!Ulcera)
 If obUlc Or Not raDT.BOF Then
  DT.fußst = auff
  FSauf = "auffällig "
  If obUlc Then FSauf = FSauf & "(Anamnesebogen:Ulcera: " & !Ulcera & ") "
  Do While Not raDT.EOF
   FSauf = FSauf & raDT!ICD & IIf(IsNull(raDT!DiagSeite), "", raDT!DiagSeite) & " (" & Format$(raDT!DiagDatum, "d.m.yy") & ") "
   raDT.Move 1
  Loop
  TabPr "Fußstatus: ", FSauf
  Dim DruckNStSeite$
  If (trp And drp) And Not (tlp And dlp) Then
   DT.NStSeite = re
   DruckNStSeite = "re (Pulse)"
  ElseIf Not (trp And drp) And (tlp And dlp) Then
   DT.NStSeite = li
   DruckNStSeite = "li (Pulse)"
  ElseIf VibrIKrep + VibrGZrep > VibrIKlip + VibrGZlip Then
   DT.NStSeite = re
   DruckNStSeite = "re (Vibr)"
  ElseIf VibrIKrep + VibrGZrep < VibrIKlip + VibrGZlip Then
   DT.NStSeite = li
   DruckNStSeite = "li (Vibr)"
  ElseIf Mfrep > Mflip Then
   DT.NStSeite = re
   DruckNStSeite = "re (Monofil.)"
  ElseIf Mfrep < Mflip Then
   DT.NStSeite = li
   DruckNStSeite = "li (Monofil.)"
  Else
   DT.NStSeite = gleich
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
   sql = "select * from diagnosen where (obdauer or (not obdauer and concat(((month(diagdatum)+2) div 3)¡ year(diagdatum)) = '" & ZQuart(Now - Verspätung) & "')) and pat_id = " & DT.Pat_id
  Else
   sql = "select * from diagnosen where (obdauer or (not obdauer and (int((month(diagdatum)+2) / 3) & year(diagdatum)) = '" & ZQuart(Now - Verspätung) & "')) and pat_id = " & DT.Pat_id
  End If
  If lies.obMySQL Then
   sql = Replace(sql, "¡", ",")
  Else
   sql = Replace(Replace(sql, "concat", ""), "¡", " & ")
  End If
  
'  Set rDT = Dtb.OpenRecordset(sql) ' Set rDT = Dtb.OpenRecordset("Diagnosen", dbOpenDynaset)
  Set raDT = Nothing
  raDT.Open sql, dbcn, adOpenDynamic, adLockReadOnly
  Dim aktDN As DFSNiveau
'  Dim maxAmp As DFSNiveau
'  Dim dt.msei$ ' re, li
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
    If aktDN > DT.maxAmp Then
     If (InStrB(raDT!DiagText, "re ") <> 0 Or Right$(raDT!DiagText, 2) = "re") Then DT.mSei = re: DT.maxAmp = aktDN Else If (InStrB(raDT!DiagText, "li ") <> 0 Or Right$(raDT!DiagText, 2) = "li") Then DT.mSei = li: DT.maxAmp = aktDN
    End If
   End If
   raDT.Move 1
  Loop
'  sql = "select * from [" & stftdb & "].jpg where pat_id = " & CStr(PID) & " and cstr(int(month(patdatum)/3)+1) & year(patdatum) = """ & ZQuart(Now - Verspätung) & """"
'  Set rafot = Dtb.OpenRecordset(sql)
  Dim raFot As New ADODB.Recordset
#If False Then
'If Forms(0).obMySQL Then
' If FtCn.State = 0 Then
'  Call acon(FotT)
' End If
' raFot.Open "select * from Jpg where pat_id = " & CStr(pid) & " and concat(floor(month(patdatum)+2) div 3,year(patdatum)) = """ & ZQuart(Now - Verspätung) & """", FtCn, adOpenDynamic, adLockReadOnly
'Else
' raFot.Open "select * from jpg where pat_id = " & CStr(pid) & " and cstr(int(month(patdatum)+2)/3) & year(patdatum) = """ & ZQuart(Now - Verspätung) & """", CStrAcc & StFtDB, adOpenDynamic, adLockReadOnly
'End If
#Else
  raFot.Open "select dokname from dokumente where pat_id = " & pid & " and dokart = 'jpg' and dokname like '%Foto%' and dokname like '% WA%'", dbcn, adOpenDynamic, adLockReadOnly
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
        If aktDN > DT.maxAmp Then DT.mSei = re
        maxDN = aktDN
        If aktWA > DT.mWA Then DT.mWA = aktWA
     ElseIf (InStrB(körperteil, "li ") <> 0 Or Right$(körperteil, 2) = "li") Then
        If aktDN > DT.maxAmp Then DT.mSei = li
        maxDN = aktDN
        If aktWA > DT.mWA Then DT.mWA = aktWA
     ElseIf (InStrB(beschreibung, "re ") <> 0 Or Right$(beschreibung, 2) = "re") Then
        If aktDN > DT.maxAmp Then DT.mSei = re
        maxDN = aktDN
        If aktWA > DT.mWA Then DT.mWA = aktWA
     ElseIf (InStrB(beschreibung, "li ") <> 0 Or Right$(beschreibung, 2) = "li") Then
        If aktDN > DT.maxAmp Then DT.mSei = li
        maxDN = aktDN
        If aktWA > DT.mWA Then DT.mWA = aktWA
     End If
    End If
   End If
   raFot.Move 1
  Loop ' While Not raFot.EOF
  
'  If dt.mwa = "" Then
  If Not raDT.BOF Then
   raDT.MoveFirst
   aktSei = unbek
   Do While Not raDT.EOF
    If InStrB(raDT!ICD, "L89.") <> 0 Then
     aktWA = Mid$(raDT!ICD, 5, 1) - 1
     If aktWA > DT.mWA Then
      DT.mWA = aktWA
      If IsNull(raDT!DiagSeite) Then
       aktSei = unbek
      Else
       Select Case UCase$(raDT!DiagSeite)
        Case "R": aktSei = re
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
  If Len(DT.mWA) = 1 Then ' 5.10.08
   Dim ArmSt$
   If DT.mWA = "0" Then
    If DT.puls = unauff Then ArmSt = "A" Else ArmSt = "C" ' 0 eher nicht entzündet
   ElseIf DT.mWA > "0" Then
    If DT.puls = unauff Then ArmSt = "B" Else ArmSt = "D"
   Else
   ' WagSt = ""
   End If
   DT.mWA = DT.mWA & ArmSt
  End If
'  End If
  Dim DruckmSei$
  If DT.mSei = unbek Then DT.mSei = aktSei
  Select Case DT.mSei
   Case re: DruckmSei = "rechts"
   Case li: DruckmSei = "links"
   Case gleich: DruckmSei = "gleich"
   Case unbek: DruckmSei = ""
  End Select
  TabPr "Schwerer betr. Fuß: ", IIf(DruckmSei = "", DruckNStSeite, DruckmSei)
  If DT.mSei = gleich Or DT.mSei = unbek Then
   DT.mSei = DT.NStSeite
   If DT.mSei = gleich Or DT.mSei = unbek Then DT.mSei = re ' meistens rechts
  End If
  If DT.mWA > "" Then
   TabPr "Grad nach Wagner + Armstrong: ", DT.mWA
  End If
  Set raFot = Nothing
'#If False Then
'  Dim posi%
'  posi = InStr(!Diagnosen, "Wagner")
'  Dim inf As Boolean, isch As Boolean, ArmSt$
'  inf = False
'  isch = False
'  If diI("T79.") Or diI("T89") Or diI("T87.4") Then inf = True
'  'If diI("M12") Or diI("M20") Or diI("M21") Or diI("M77.3") Or diI("M84") Or diI("M84.17") Then def = True
'  If (trp And drp) Or (tlp And dlp) Or (trp And dlp) Or (tlp And drp) Then isch = True
'  ArmSt = "A"
'  If inf Then
'   ArmSt = "B"
'   If isch Then ArmSt = "D"
'  ElseIf isch Then
'   ArmSt = "C"
'  End If
'  If posi <> 0 Then
'   TabPr "Grad nach Wagner und Armstrong: ", IIf(posi <> 0, Mid$(!Diagnosen, posi + 7, 1) + ArmSt, "0A")
'  End If
'#End If
  Set raDT = Nothing
  Call raDT.Open("select * from diagnosen where pat_id = " & pid & " and icd like '" & "M14.6%" & "' and diagsicherheit in (""G"",""V"")", dbcn, adOpenDynamic, adLockReadOnly)
  If Not raDT.BOF Then
   Dim Seite$
   If Not IsNull(raDT!DiagSeite) Then
    Seite = raDT!DiagSeite
    Select Case raDT!DiagSeite
     Case "R": DT.oap = re
     Case "L": DT.oap = li
     Case "B": DT.oap = gleich ' heißt hier: unbekannt
     Case Else: DT.oap = unbek ' heißt hier: keine
    End Select
   Else
    Seite = "ja"
    DT.oap = gleich
   End If
   TabPr "Osteoarthropathie ", Seite
  Else
   DT.oap = unbek ' heißt hier: keine
   TabPr "Osteoarthropathie: ", "keine"
  End If
  Set raDT = Nothing
  
  raKH.Open "select * from kheinweis where pat_id = " & pid & " and zeitpunkt > " & datform(Now - 100), dbcn, adOpenDynamic, adLockReadOnly
  If Not raKH.EOF Then
   DT.khew = True
   TabPr "Krankenhauseinweisung", "(am " + Format$(raKH!Zeitpunkt, "dd/mm/yy") + IIf(tfeld(raKH!Ziel) <> "nkenhauseinweisung", " nach: " + raKH!Ziel, "") + _
             " mit: " + tfeld(raKH!Diagnose) + ")"
  End If
 ElseIf tfeld(!Ulcera) = "" And tfeld(![Puls Atp]) = "" And tfeld(![Puls Adp]) = "" And tfeld(![Vibration IK]) = "" And tfeld(![Vibration Großzehe]) = "" And tfeld(![Kalt-Warm]) = "" Then
  DT.fußst = ndok
  TabPr "Fußstatus: ", "" '"nicht dokumentiert"
 Else ' obUlc Or Not raDT.BOF Then
  DT.fußst = unauff
  TabPr "Fußstatus: ", "unauffällig"
 End If ' obUlc Or Not raDT.BOF Then
 
 If DT.dtyp = "1" Then
  If DT.PrRR <> "" Then
   TabPr "Blutdruck: ", DT.PrRR
  End If
  TabPr "HbA1c: ", DMPHbA1c
  TabPr "Serum-Kreatinin: ", DMPCrea
  TabPr "Urin-Albumin: ", DMPUAlb
 End If
 
' Dim oblaser As Boolean
 Dim kontrah$
 DT.oblaser = True
 If Not obPosi(![Netzhaut gelasert]) Then
  DT.oblaser = False
'  rEi.Seek "=", PID
  Set raEi = Nothing
  Call raEi.Open("select * from eintraege where pat_id = " & pid, dbcn, adOpenDynamic, adLockReadOnly)
  If Not raEi.BOF Then
   Do
    Dim StartPosi&
    If raEi.EOF Then Exit Do
    kontrah = Replace(raEi!Inhalt, " ", "")
    StartPosi = 1
    StartPosi = InStr(kontrah, "Netzhautschongelasert")
    If StartPosi > 1 Then StartPosi = StartPosi + 20 ' ca.
    If StartPosi = 0 Then StartPosi = 1
    If InStrB(StartPosi, kontrah, "laser") <> 0 Then
     DT.oblaser = True
     Exit Do
    End If
    raEi.Move 1
   Loop
  End If
 End If
' Handlungsbedarf: spätere Lasertherapie
 If DT.dtyp = "2" Then TabPr "Lasertherapie: ", IIf(DT.oblaser, "ja", "nein")
 
 DT.hypoZAn = 0
 If obPosi(![Fremde Hilfe pa]) Then
  If IsNumeric(![Fremde Hilfe pa]) Then
   DT.hypoZAn = ![Fremde Hilfe pa]
  End If
  TabPr "Schw.Hypoglyk./12 Mon: ", CStr(![Fremde Hilfe pa])
 Else
  TabPr "Schw.Hypoglyk./12 Mon: ", IIf(tfeld(![Fremde Hilfe pa]) = "", "keine Angaben", "keine")
 End If
 rs.Open "select count(0) as ct from eintraege where pat_id = " & pid & " and art = 'hypo' and zeitpunkt >= " & datform(QAnf(ZQuart(Now() - Verspätung))), dbcn, adOpenStatic, adLockReadOnly
 DT.hypoZKK = rs!ct
 Set rs = Nothing
 rs.Open "select count(0) as ct from eintraege where pat_id = " & pid & " and art = 'kra' and zeitpunkt >= " & datform(QAnf(ZQuart(Now() - Verspätung))), dbcn, adOpenStatic, adLockReadOnly
 DT.krZKK = rs!ct
 Set rs = Nothing
 
' Handlungsbedarf: Krankenhauseinweisungen
 DT.krZAn = 0
 If obPosi(![keto pa]) Then
  If IsNumeric(![keto pa]) Then
   DT.krZAn = ![keto pa]
  End If
  TabPr "Krankenhausaufenthalte / Ketoazidosen /12 Mon: ", CStr(![keto pa])
 Else
  TabPr "Krankenhausaufenthalte / Ketoazidosen /12 Mon: ", IIf(tfeld(![Fremde Hilfe pa]) = "", "keine Angaben", "keine")
 End If
 
 If diI("Z49", , QAnf(ZQuart(Now() - Verspätung))) Then
  DT.neuDial = True
  Set raDT = Nothing
 End If
 If diI("H54", , QAnf(ZQuart(Now() - Verspätung)), , "diagseite in (' ','B')") Then
  DT.neuErbl = True
  Set raDT = Nothing
 End If
 If diI("Z44", , QAnf(ZQuart(Now() - Verspätung)), True) Then
  DT.neuAmp = True
  Set raDT = Nothing
 End If
 If diI("I21", , QAnf(ZQuart(Now() - Verspätung)), True) Then
  DT.neuMI = True
  Set raDT = Nothing
 End If
 
 If DT.dtyp = "1" Then TabPr "Lasertherapie: ", IIf(DT.oblaser, "ja", "nein")
 
 Dim metText$
 DT.obNI = (DT.Crea > 1.3)
 
 
 DmPStr = DmPStr & vbCrLf
 If DT.dtyp <> "1" Then
  TabPr "Glibenclamid: ", IIf(DT.obGlib, IIf(DT.VorM(0) = 0 And DT.obVorb, "(unverändert) ", "  ") & "ja", IIf(DT.VorM(0) > 0, "(unverändert) ", "  ") & IIf(DT.obNI, "Kontraindikation", "Nein"))
  If Not DT.obGlib And DT.obNI Then DT.obGlib = adki
  If DT.obmetf Then
   metText = "  Ja"
   If DT.VorM(1) = 0 And DT.obVorb Then metText = "(unverändert)" & metText
   If DT.obNI Then metText = metText + ": Sollte evtl. aufgrund Creatininerhöhung abgesetzt werden"
  Else
   metText = IIf(DT.obNI, " Kontraindikation", " Nein")
   If DT.VorM(1) = 1 Or DT.VorM(1) = 2 Then metText = "(unverändert)" & metText
  End If
  If Not DT.obmetf And DT.obNI Then DT.obmetf = adki
  TabPr "Metformin: ", metText
  TabPr "Glucosidase-Inhib.: ", IIf(DT.obGlucI, IIf(DT.VorM(2) = 0 And DT.obVorb, "(unverändert) ", "  ") & "Ja", IIf(DT.VorM(2) > 0, "(unverändert) ", "  ") & "Nein")
  TabPr "Sonst. Sulf'hst./Glinide: ", IIf(DT.obSHGlin, IIf(DT.VorM(3) = 0 And DT.obVorb, "(unverändert) ", "  ") & "Ja", IIf(DT.VorM(3) > 0, "(unverändert) ", "  ") & "Nein")
  If Not DT.obGlit And diI("I50") Then DT.obGlit = adki
  TabPr "Glitazone: ", IIf(DT.obGlit = adja, IIf(DT.VorM(4) = 0 And DT.obVorb, "(unverändert) ", "  ") & "Ja", IIf(DT.VorM(4) > 0, "(unverändert) ", "  ") & IIf(DT.obGlit = adki, "Kontraindikation", "Nein"))
  TabPr "Sonstige Antidiab.: ", IIf(DT.obSonstAD = adja, IIf(DT.VorM(12) = 0 And DT.obVorb, "(unverändert) ", "  ") & "Ja", IIf(DT.VorM(12) > 0, "(unverändert) ", "  ") & "Nein")
 End If ' dt.dtyp <> "1"
 Dim TherArtAkt$
' TherArtAkt = TherArtAkt
 TherArtAkt = TherArt(DT.Pat_id, 0)
 DT.tart = TherArtAkt
' Select Case TherArt(DT.Pat_id, 0)
'  Case "Diät?":  DT.tart = diät
'  Case "": DT.tart = unbek
'  Case "(I)CT?", "I/CT?", "ICT": DT.tart = ict
'  Case "CSII":   DT.tart = CSII
'  Case "CT":     DT.tart = ct
'  Case "Komb":   DT.tart = komb
'  Case "OAD":     DT.tart = oad
' End Select
' If dt.tart = csii Then Stop
 TabPr "Insulin:  ", IIf(DT.obIns, IIf(DT.insz > 2, IIf(DT.VorM(5) = 1, "(unverändert) ", "  ") & "ja, intensiviert", IIf(DT.VorM(5) = 2, "(unverändert) ", "  ") & "ja, " & TherArtAkt), IIf(DT.VorM(5) = 0 And DT.obVorb, "(unverändert) ", "  ") & "Nein")
 TabPr "Insulin-Analoga: ", IIf(DT.obAnal, IIf(DT.insz > 2, IIf(DT.VorM(6) = 1, "(unverändert) ", "  ") & "ja, intensiviert", IIf(DT.VorM(6) = 2, "(unverändert) ", "  ") & "ja, " & TherArtAkt), IIf(DT.VorM(6) = 0 And DT.obVorb, "(unverändert) ", "  ") & "Nein")
 DmPStr = DmPStr & vbCrLf
 TabPr "HMG-CoA-Red'-Hemmer: ", IIf(DT.obHMG, IIf(DT.VorM(7) = 1, "(unverändert) ", "  ") & "ja", IIf(DT.VorM(7) = 0 And DT.obVorb, "(unverändert) ", "  ") & "nein")
 TabPr "Antihypertensiva: ", IIf(DT.obAntihyp, IIf(DT.VorM(8) = 1, "(unverändert) ", "  ") & "ja", IIf(DT.VorM(8) = 0 And DT.obVorb, "(unverändert) ", "  ") & "nein")
 TabPr "Betablocker: ", IIf(DT.obBetabl, IIf(DT.VorM(10) = 1, "(unverändert) ", "  ") & "ja", IIf(DT.VorM(10) = 0 And DT.obVorb, "(unverändert) ", "  ") & "nein")
 TabPr "ACE-Hemmer: ", IIf(DT.obACEH, IIf(DT.VorM(11) = 1, "(unverändert) ", "  ") & "ja", IIf(DT.VorM(11) = 0 And DT.obVorb, "(unverändert) ", "  ") & "nein")
 TabPr "Thrombozytenhemmer: ", IIf(DT.obThro, IIf(DT.VorM(9) = 1, "(unverändert) ", "  ") & "ja", IIf(DT.VorM(9) = 0 And DT.obVorb, "(unverändert) ", "  ") & "nein")

 DmPStr = DmPStr & vbCrLf
 TabPr " Schulungen vor DMP:", IIf(DT.SW(0) > 0, " D.m.: " & Format$(DT.SW(0), "d.m.yy"), "") & IIf(DT.SW(1) > 0, " Hypert: " & Format$(DT.SW(1), "d.m.yy"), "") & IIf(DT.SW(2) > 0, " keine: " & Format$(DT.SW(2), "d.m.yy"), "")
' Schulungen empfohlen
 Set raLau = Nothing
 raLau.Open "select zeitpunkt,wert from (" & lsql & ") as innen where abkü = ""HBA1C"" order by zeitpunkt desc", dbcn, adOpenDynamic, adLockReadOnly
 Dim zpd1$, zpdl$, zpr1$, zprl$
 If Not raLau.BOF Then
  If raLau!Wert >= 8# Then
   DMSchulz = SchulzBest(!Pat_id, zpd1, zpdl, Now() - 2 * 365)
   If ((DT.tart = oad Or DT.tart = offen Or DT.tart = diät) And DMSchulz < 4) Or DMSchulz < 6 Then
'    Debug.Print Pat_id
    DT.SE(0) = Now
   End If
  End If
 End If
 TabPr " Schulungen empfohl:", IIf(DT.SE(0) > 0, " D.m.: " & Format$(DT.SE(0), "d.m.yy"), "") & IIf(DT.SE(1) > 0, " Hypert: " & Format$(DT.SE(1), "d.m.yy"), "") & IIf(DT.SE(2) > 0, " keine: " & Format$(DT.SE(2), "d.m.yy"), "")
 TabPr " DMP-Schulung D.m.:", IIf(DT.ds(0) > 0, " ja: " & Format$(DT.ds(0), "d.m.yy"), "") & IIf(DT.ds(1) > 0, " nein: " & Format$(DT.ds(1), "d.m.yy"), "") & IIf(DT.ds(2) > 0, " nicht mögl: " & Format$(DT.ds(2), "d.m.yy"), "")
 TabPr " DMP-Schulung Hypt:", IIf(DT.HS(0) > 0, " ja: " & Format$(DT.HS(0), "d.m.yy"), "") & IIf(DT.HS(1) > 0, " nein: " & Format$(DT.HS(1), "d.m.yy"), "") & IIf(DT.HS(2) > 0, " nicht mögl: " & Format$(DT.HS(2), "d.m.yy"), "")
 If DT.kSE <> 0 Then TabPr " Zuletzt k.Sch.empf: ", Format$(DT.kSE, "d.m.yy")
  
 If obPosi(![Jahr letzte Diabetesschulung]) Then
  SchulStr = ![Jahr letzte Diabetesschulung] + IIf(IsNull(![Ort Schulung]), "", " (" + ![Ort Schulung] + ")")
 Else
  SchulStr = "keine"
 End If
 TabPr "Schulung vor Zuweisung:", SchulStr
 Dim sqls$
 sqls$ = _
 "SELECT leistungen.pat_id as pat_id, leistungen.zeitpunkt as zeitpunkt, ucase(leistungen.leistung) as lst" & _
 " FROM leistungen INNER JOIN faelle ON faelle.pat_id = leistungen.pat_id AND faelle.quartal = cstr(int(month(leistungen.zeitpunkt)+2) / 3) & year(leistungen.zeitpunkt) " & _
 " where faelle.pat_id = " & CStr(pid)
 If lies.obMySQL Then sqls = Replace(Replace(Replace(sqls, "int(", "floor("), "cstr", ""), " / ", " div ")
' Set lpp = Dtb.OpenRecordset(sqls, dbOpenDynaset)
 Call lapp.Open(sqls, dbcn, adOpenDynamic, adLockReadOnly)
 DMSchulz = 0
 HSchulz = 0
 zpd1 = Format$(Now + 1, "d.m.yy")
 zpdl = "1.1.1900"
 zpr1 = Format$(Now + 1, "d.m.yy")
 zprl = "1.1.1900"
 If Not lapp.BOF Then
  lapp.MoveFirst
  Do While Not lapp.EOF
   Call SchulungszifferZuord(lapp!lst, DMSchulz, HSchulz, lapp!Zeitpunkt, zpd1, zpdl, zpr1, zprl)
   lapp.Move 1
  Loop
 End If
 
 DMSchulz = SchulzBest(!Pat_id, zpd1, zpdl)
 SchulStr = IIf(DMSchulz > 1, CStr(DMSchulz) + " (Diabetes: " + zpd1 + "-" + zpdl + ")", "") + IIf(HSchulz > 1, CStr(HSchulz) + " (Hypertonie: " + zpr1 + "-" + zprl + ")", "")
 If SchulStr = "" Then SchulStr = "keine"
 TabPr "Bish.Schulungen bei mir:", SchulStr
 DmPStr = DmPStr & vbCrLf
 If DT.dtyp = "2" Then
  TabPr "Aufgabe Tabak empf.:", IIf(DT.Tabak, "ja", "nein")
  If !Größe <> 0 Then
''  TabPr "Ernährungsber. empf.:", IIf(CDbl(Replace(![bmi], ".", ",")) > 24.9, "ja", "nein")
'   Dim bmi!
'   bmi = !Gewicht * IIf(!Gewicht < 3, 100, 1) / !Größe / !Größe '* 10000
'   Do
'    If bmi = 0 Then Exit Do
'    If bmi > 8 Then Exit Do
'    bmi = bmi * 10
'   Loop
'   If bmi > 24.9 Then dt.ernb = True
   If DT.bmi >= 25 Then DT.ernb = True
   TabPr "Ernährungsber. empf.:", IIf(DT.ernb, "ja", "nein")
  End If
 End If
 DT.hbEmpf = IIf(DT.bekHb > 6.5, senken, halten)
 TabPr "Zielvereinb. HbA1c:", IIf(DT.hbEmpf = senken, "senken", "halten")
 Dim rrEmpfehlung$
 rrEmpfehlung = rrEmpf(DT.RRsyst, DT.RRdiast, pid)
 DT.rrEmpf = IIf(rrEmpfehlung = "halten", halten, senken)
 TabPr "Zielvereinb. Blutdruck:", rrEmpfehlung ' manuell auszufüllen
 Dim AugU$, AugUDat As Date
 Dim rAdo As New ADODB.Recordset
 AugU = ""
' Set rdo = Dtb.OpenRecordset("select * from [" + QMdbAkt + "].Dokumente where pat_id = " & CStr(Pat_id) + " and zeitpunkt > now() - 550 order by zeitpunkt desc", dbOpenDynaset) ' 1,5 Jahre
 rAdo.Open "select * from dokumente where pat_id = " & CStr(DT.Pat_id) & " and zeitpunkt > " & datform(Now() - 550) & " order by zeitpunkt desc", dbcn, adOpenDynamic, adLockReadOnly
 If Not rAdo.BOF Then
  Do While Not rAdo.EOF
   If InStrB(LCase(rAdo!DokName), "augen") <> 0 Then
    AugUDat = DatInStr(rAdo!DokName, Year(rAdo!Zeitpunkt))
    AugU = "Befund vo" + IIf(AugUDat = CDate(0), "n " + Format$(rAdo!Zeitpunkt, "yyyy"), "m " + Format$(AugUDat, "dd/mm/yy")) + " vorliegend"
    If AugUDat = CDate(0) Then AugUDat = rAdo!Zeitpunkt ' dann war die letzte Untersuchung vermutlich frühestens am 1.1. des Scan-Jahres
    Exit Do
   End If
   rAdo.Move 1
  Loop
 End If
' Dim rEin As DAO.Recordset
 If lies.obMySQL Then
  sql = "select * from eintraege where pat_id = " & CStr(pid) & " and ( art = 'aug' or (" + artspez + " and ((inhalt like ""%augenb%"" and not inhalt like ""%augenbl%"" and not inhalt like ""%augen"") or (inhalt like ""%augenarzt%"" or inhalt like ""%augenärzt%"") or (inhalt like ""% aa%"" and not inhalt like ""% aag%""))) or (art = ""aa"" or art = ""augen"")))" & " and zeitpunkt > " & datform(Now() - 550) & " order by zeitpunkt desc limit 3"
 Else
  sql = "select top 3 * from eintraege where pat_id = " & CStr(pid) & " and ( art = 'aug' or (" + artspez + " and (inhalt like ""%augenb[!l]%"" or inhalt like ""%augen[aä]rzt%"" or inhalt like ""% aa[!g]%"")) or (art = ""aa"" or art = ""augen"")))" + " and zeitpunkt > " & datform(Now() - 550) & " order by zeitpunkt desc"
 End If
' Set rEin = Dtb.OpenRecordset(sql)
 Set raEi = Nothing
 raEi.Open sql, dbcn, adOpenDynamic, adLockReadOnly
 Do While Not raEi.EOF
  AugU = AugU & vbCrLf & vbTab & vbTab & "Eintrag am " & Format$(raEi!Zeitpunkt, "d/m/yy") & ":" + vbTab + raEi!Inhalt
  raEi.Move 1
 Loop
 Dim uebw As New ADODB.Recordset 'DAO.Recordset
  fiabfr = "SELECT Pat_ID, FID, Form_ID, ZeitPunkt, Nr, FeldNr, Feld, FeldInh, form_abk FROM (((forminhfeld LEFT JOIN forminhkopf ON forminhfeld.foid=forminhkopf.foid) LEFT JOIN formulare ON formulare.formid=forminhkopf.form_id) LEFT JOIN forminhaltfeld ON forminhfeld.feldvw=forminhaltfeld.feldvw) LEFT JOIN forminhaltfeldinh ON forminhfeld.feldinhvw=forminhaltfeldinh.feldinhvw "
  Call uebw.Open(fiabfr & " where pat_id = " & CStr(pid) + " and form_abk = ""uew"" and feld = ""Ueberweisung_an"" and feldinh = ""Augenheilkunde"" and zeitpunkt > " & datform(Now() - 550) & " order by zeitpunkt desc", dbcn, adOpenDynamic, adLockReadOnly)
' Set uebw = Dtb.OpenRecordset(, dbOpenDynaset)
 If Not uebw.BOF Then
  If AugUDat = CDate(0) Or uebw!Zeitpunkt > AugUDat Then
   AugU = AugU + IIf(AugU = "", "", ", ") + "Untersuchung veranlasst am " + Format$(uebw!Zeitpunkt, "dd/mm/yy")
  End If
 End If
 If AugU = "" And tfeld(![Augensp zuletzt]) <> "" Then
  AugU = ![Augensp zuletzt]
  If tfeld(![Augensp Befund]) <> "" Then
   AugU = AugU + " (Befund: " + ![Augensp Befund] + ")"
  End If
 End If
 If AugUDat > Now - 500 Then DT.aug = durchg Else DT.aug = veranl
 TabPr "Augenuntersuchung:", AugU
 TabPr "Empf. Dok'intervall:", IIf(DT.bekHb <= 6.5 And (Not obPosi(tfeld(![Unterzucker pM])) Or InStrB(tfeld(![Unterzucker pM]), "selten") <> 0), "viertel- oder halbjährlich", "vierteljährlich")
#If False Then
#End If
End With ' raana
DMPString = DmPStr
lDMPPat_id = pid
üdt = DT
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DMPString/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
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
  ursp = Trim$(Replace(Replace(Replace(ursp, "ca.", ""), "ca", ""), "seit", ""))
  pos = InStr(ursp, "/")
  If pos > 0 Then
   ursp = Mid$(ursp, pos + 1)
   buch = Mid$(ursp, 3)
   If Not IsNumeric(buch) Then
    buch = Left(ursp, 1)
    If buch > 5 Then
     ursp = "19" & ursp
    Else
     ursp = "20" & ursp
    End If
   End If
  End If
  pos = InStr(ursp, " ")
  If pos > 0 Then
   If IsNumeric(Left(ursp, 1)) Then
    ursp = Left(ursp, pos - 1)
   Else
    ursp = Mid$(ursp, pos + 1)
   End If
  End If
  pos = InStr(ursp, "-")
  If pos > 0 Then
   ursp = Left(ursp, pos - 1)
  End If
  pos = InStr(ursp, "a")
  If pos > 0 Then
   ursp = Left(ursp, pos - 1)
  End If
  pos = InStr(ursp, "J")
  If pos > 0 Then
   ursp = Left(ursp, pos - 1)
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
Function TherUmw$(TherArt As TherapieArt)
 Select Case TherArt
  Case diät: TherUmw = "Diät?"
  Case offen: TherUmw = ""
  Case ict: TherUmw = "ICT"
  Case csii: TherUmw = "CSII"
  Case ct: TherUmw = "CT"
  Case komb: TherUmw = "Komb"
  Case oad: TherUmw = "OAD"
  Case Else
    MsgBox "Fehler bei Therapieart: " & TherArt
    Tüt 1760, 1000
    Stop
 End Select
' If dt.tart = csii Then Stop

End Function
Function ImportFolderHerricht()
 Dim fld As Folder, Fil As File, FPath$
 On Error GoTo fehler
 Call VerzPrüf(hVerz)
 Set fld = FSO.GetFolder(hVerz)
 For Each Fil In fld.Files
  If Right$(Fil.Name, 4) = ".BDT" Then
   FPath = Fil.path
   On Error Resume Next
   Name FPath As Replace(FPath, ".BDT", ".txt")
   Kill FPath
   On Error GoTo fehler
  End If
 Next Fil
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ImportFolderHerricht/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ImportFolderHerricht()
'Function LeistungsExport0()
' Dim op$, BDT As New BDTSchreib
' On Error GoTo fehler
'' Call ImportFolderHerricht
'' LeistBDT = hVerz + "LEIST " + Format$(Now, "dd/mm/yy HH.MM") + ".BDT"
'' Open LeistBDT For Output As #310
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
' AnwPfad = CurrentDb.Name
'#Else
' AnwPfad = App.path
'#End If
'Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LeistungsExport0/" + AnwPfad)
' Case vbAbort: Call MsgBox("Höre auf"): End
' Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
' Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
'End Select
'End Function ' LeistungsExport0
Function LeistungsExport1(BDT As BDTSchreib, Pat_id&, Leist$, Datu As Date, Optional QUZeit$, Optional nurkasse%)
' Dim q As DAO.Recordset
 Dim rFa As New ADODB.Recordset, rNa As New ADODB.Recordset
 Dim op$
 On Error GoTo fehler
' Set rFa = TabÖff("faelle", "Auswahl")
 rFa.Open "select * from faelle where pat_id = " & Pat_id & " and bhfe1 >= " & datform(Datu) & IIf(nurkasse, " and schgr <> '90' ", "") & " order by bhfe1", dbcn, adOpenStatic, adLockReadOnly
 If rFa.BOF Then ' wenn kein Fall in die Zeitschiene paßt, dann den letzten nehmen
  Set rFa = Nothing
  rFa.Open "select * from faelle where pat_id = " & Pat_id & IIf(nurkasse, " and schgr <> '90' ", "") & " order by bhfe1 desc", dbcn, adOpenStatic, adLockReadOnly
 End If
 If rFa.BOF Then Exit Function
 rNa.Open "select * from namen where pat_id = " & Pat_id, dbcn, adOpenStatic, adLockReadOnly
 Forms(0).Ausgabe = "Trage Leistung " & Leist & " für Pat_ID " & Pat_id & " (" & rNa!NachName & ", " & rNa!Vorname & ") mit Datum " & Format$(Datu, "dd.mm.yyyy") & " ein." & vbCrLf & altAusgabe
 LeistungsExport1 = 1
 altAusgabe = Forms(0).Ausgabe
 Call BDT.Satzart(IIf(rFa!SchGr = "90", "0190", "0102")) ' Satzidentifikation
' bei 0101 entstehen bei zwei Aufrufen fehlerfrei zwei neue Kassenfaelle, jeder mit der Leistung
' bei 190 entsteht ein neuer Privatfall, bei 6100 läuft alles ohne Fehler durch, aber keine Leistung steht drin
' bei 6200 entsteht ein neuer Kassenfall
'     op = format$(3 + 4 + 4, "000") + "8000" + CStr(f!s8000)
'     Print #310, zsu(op)
'     op = format$(3 + 4 + 5, "000") + "8100" + CStr(f!s8100)
'     Print #310, zsu(op)
 Call BDT.PatID(rFa!Pat_id)
#If False Then
     If 1 = 0 Then 'Auswirkung bisher nicht geprüft 31.7.05 (3x)
      op = Format$(3 + 4 + Len(n!NVorsatz), "000") + "3100" + CStr(n!NVorsatz)
      Print #310, ZSU(op)
     End If
     op = Format$(3 + 4 + Len(n!NachName), "000") + "3101" + CStr(n!NachName)
     Print #310, ZSU(op)
     op = Format$(3 + 4 + Len(n!Vorname), "000") + "3102" + CStr(n!Vorname)
     Print #310, ZSU(op)
     op = Format$(3 + 4 + 8, "000") + "3103" + Format$(n!gebdat, "ddmmyyyy")
     Print #310, ZSU(op)
     op = Format$(3 + 4 + Len(IIf(IsNull(n!Versichertennummer), "", n!Versichertennummer)), "000") + "3105" + IIf(IsNull(n!Versichertennummer), "", n!Versichertennummer)
     Print #310, ZSU(op)
'     op = format$(3 + 4 + Len(n!Straße), "000") + "3107" + n!Straße
'     Print #310, ZSU(op)
'     op = format$(3 + 4 + Len(n!Plz), "000") + "3112" + n!Plz
'     Print #310, ZSU(op)
'     op = format$(3 + 4 + Len(n!Ort), "000") + "3113" + n!Ort
'     Print #310, ZSU(op)
''     GoTo leistungdirekt
     If Not IsNull(n![KVKStatus]) And n![KVKStatus] > 0 And Trim$(n![KVKStatus]) <> "" Then
      op = Format$(3 + 4 + Len(n![KVKStatus]), "000") + "3108" + n![KVKStatus]
      Print #310, ZSU(op)
     End If
     op = Format$(3 + 4 + 1, "000") + "3110" + IIf(n!Geschlecht = "m", "1", IIf(n!Geschlecht = "w", "2", ""))
     Print #310, ZSU(op)
#End If
#If False Then
     If q!SchGr <> "90" Then
      op = Format$(3 + 4 + 5, "000") + "4101" + Quartal
      Print #310, ZSU(op)
      op = Format$(3 + 4 + 8, "000") + "4102" + Format$(rFa!ausgst, "ddmmyyyy")
      Print #310, ZSU(op)
      op = Format$(3 + 4 + Len(rFa!VKNr), "000") + "4104" + rFa!VKNr
      Print #310, ZSU(op)
     End If
#End If
     If Not IsNull(rFa!KtrAbrB) Then  ' bei Privaten
      BDT.SAdd "4106", rFa!KtrAbrB
     End If
     If Not IsNull(rFa!AbrAr) Then  ' bei Privaten
      BDT.SAdd "4107", rFa!AbrAr
     End If
     If rFa!SchGr <> "90" Then
      BDT.DAdd "4109", rFa!lVorl
      BDT.TAdd "4110", rFa!lVorl
      BDT.SAdd "4111", rFa!IK
      If Not IsNull(rFa!KVKs) Then ' bei Pat_id 43
       BDT.SAdd "4112", rFa!KVKs
      End If
      If Not IsNull(rFa!KVKserg) Then
       BDT.SAdd "4113", rFa!KVKserg
      End If
     End If
     If Not IsNull(rFa!GebOr) Then ' bei Privaten
      BDT.SAdd "4121", rFa!GebOr
     End If
     If Not IsNull(rFa!AbrGb) Then ' bei Privaten
      BDT.SAdd "4122", rFa!AbrGb
     End If
     BDT.SAdd "4144", "TM#" + IIf(IsNull(rFa!TMFNr), Space$(11), rFa!TMFNr)
     BDT.DAdd "4150", rFa!BhFB
     BDT.DAdd "4151", rFa!BhFE1
'     op = format$(3 + 4 + 8, "000") + "4152" + IIf(rFa!BhFE2 = 0, "00000000", format$(rFa!BhFE2, "ddmmyyyy"))
'     Print #310, ZSU(op)
     Dim Üw$
     Üw = IIf(IsNull(rFa!ÜbwV), "", rFa!ÜbwV)
     If Üw <> "" Then
      BDT.SAdd "4218", Üw
     End If
     Üw = IIf(IsNull(rFa!AndÜw), "", rFa!AndÜw)
     If Üw <> "" Then
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
     BDT.SAdd "5001", Leist
 Close #310
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LeistungsExport/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LeistungsExport1
Function doPorto()
 Const Verz$ = "p:\gefaxt\"
 Dim BDT As New BDTSchreib
 Dim rs As New ADODB.Recordset, Zahl&
 Dim DateiDatum As Date
 Dim Datei$, Pidpos&, pidp2%, pid&, Datum As Date
 On Error GoTo fehler
' If DBCn.State = 0 Then
'  Call DBVerb.cnVorb("", "anamnesebogen", "Patientendaten")
'  Set DBCn = DBVerb.wCn
'  LVobMySQL = InStrB(UCase$(DBVerb.CnStr), "MYSQL") > 0
' End If
 If Not BDT.Start(hVerz, "Leist") Then
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
  If InStrB(Datei, " PID ") <> 0 Then
   DateiDatum = FileDateTime(Verz & Datei)
   If DateiDatum >= QAnf(ZQuart(Now() - Verspätung)) Then
    Pidpos = InStr(Datei, " PID ") + 5
    pidp2 = InStr(Pidpos, Datei, " ")
    pid = Mid$(Datei, Pidpos, pidp2 - Pidpos)
    If InStrB(Datei, " DMP-Daten ") <> 0 Then
     Pidpos = InStr(Datei, "DMP-Daten ") + 14 ' ... vom
     pidp2 = InStr(Pidpos, Datei, " ")
     Datum = Mid$(Datei, Pidpos, pidp2 - Pidpos)
'    Call LeistungsExport1(pid, "01600", datum, CDate("18:00"))
     Zahl = Zahl + LeistungsExport1(BDT, pid, "40120", Datum, CDate("18:00"), True)
    ElseIf InStrB(Datei, "Arztbrief vom") <> 0 Then
     Pidpos = InStr(Datei, "Arztbrief vom ") + 14
     pidp2 = InStr(Pidpos, Datei, " ")
     Datum = Mid$(Datei, Pidpos, pidp2 - Pidpos)
'    Call LeistungsExport1(pid, "01601", datum, CDate("18:00"))
     Zahl = Zahl + LeistungsExport1(BDT, pid, "40120", Datum, CDate("18:00"), True)
    End If
   End If
  End If
  Datei = Dir
 Loop
 
 rs.Open "select b.* from (select pat_id, date(zeitpunkt) as datum, time(zeitpunkt) as zeit from briefe b where name like '%brief%' and zeitpunkt between " & datform(QAnf(ZQuart(Now() - Verspätung))) & " and " & datform(QEnd(ZQuart(Now() - Verspätung))) & " group by pat_id, date(zeitpunkt)) as b left join leistungen l on b.pat_id = l.pat_id  and leistung like '4012%' where isnull(leistung)", dbcn, adOpenStatic, adLockReadOnly
 Do While Not rs.EOF
  Zahl = Zahl + LeistungsExport1(BDT, rs!Pat_id, "40120", rs!Datum, CDate(rs!zeit), True)
  rs.Move 1
 Loop
 BDT.Schreib
' Close #310
 MsgBox "Datei '" & BDT.z & "' neu mit " & Zahl & " Leistungen zu den Briefen in '" & Verz & "' und zu Arztbriefen in der Tabelle `briefe` für Kassenpatienten erstellt!"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doPorto/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Function tuBriefeLeiDok(frm As Lese)
 On Error GoTo fehler
 Dim sql2$
 Dim rF As New ADODB.Recordset, rB As New ADODB.Recordset, rl As New ADODB.Recordset
 sql1 = "select min(bhfb) as bfb, min(schgr) as schgr, pat_id, min(fid) as fid from faelle where schgr in ('21','23','24','00') and quartal = """ & ZQuart(Now - Verspätung) & """ group by pat_id"
' sql1 = "select f.fid as fid, date(b.zeitpunkt) as tag, b.* from briefe b left join faelle f on b.pat_id = f.pat_id where quartal = '" & ZQuart(Now - Verspätung) & "' and schgr in ('21','23','24','00') and b.zeitpunkt >= " & datform(QAnf(ZQuart(Now() - 20))) & " and name like '%.doc' and (name like '%brief%' or name like '%dmp-daten%' or name like '%nachricht an%') order by b.pat_id,b.zeitpunkt;"
' Call LeistungsExport0
 rF.Open sql1, dbcn, adOpenStatic, adLockReadOnly
 Do While Not rF.EOF
  If rF!Pat_id = 2155 Then Stop
  sql1 = "select * from briefe where pat_id = " & rF!Pat_id & " and zeitpunkt >= " & datform(QAnf(ZQuart(Now() - 20))) & " and name like '%.doc' and (name like '%brief%' or name like '%nachricht an%') and zeitpunkt >= " & datform(#10/16/2007#)
  Set rB = Nothing
  rB.Open sql1, dbcn, adOpenDynamic, adLockReadOnly
  If Not rB.EOF Then
   Set rl = Nothing
   rl.Open "select * from leistungen where fid = " & rF!FID & " and leistung in ('01601')", dbcn, adOpenStatic, adLockReadOnly
   If rl.BOF Then
'    Call LeistungsExport1(rF!Pat_id, "01601", min(DateValue(rB!Zeitpunkt), #12/31/2007#), CDate("18:00"))
    Call LeistungsExport1(rF!Pat_id, "40120", min(DateValue(rB!Zeitpunkt), #12/31/2007#), CDate("18:00"))
   End If
  Else
   sql1 = "select * from briefe where pat_id = " & rF!Pat_id & " and zeitpunkt >= " & datform(QAnf(ZQuart(Now() - 20))) & " and name like '%.doc' and (name like '%dmp-daten%') and zeitpunkt >= " & datform(#10/16/2007#)
   Set rB = Nothing
   rB.Open sql1, dbcn, adOpenStatic, adLockReadOnly
   If Not rB.EOF Then
    Set rl = Nothing
    rl.Open "select * from leistungen where fid = " & rF!FID & " and leistung in ('01600')", dbcn, adOpenStatic, adLockReadOnly
    If rl.BOF Then
'     Call LeistungsExport1(rF!Pat_id, "01600", min(DateValue(rB!Zeitpunkt), #12/31/2007#), CDate("18:00"))
     Call LeistungsExport1(rF!Pat_id, "40120", min(DateValue(rB!Zeitpunkt), #12/31/2007#), CDate("18:00"))
    End If
   End If
  End If
  
  rF.Move 1
 Loop
 MsgBox "Fertig meit Leistungsdoku"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in tuBriefeLeiDok/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Function alleDMPLeiDok(frm As Lese)
 Dim DefDB$
 Dim rsA As New ADODB.Recordset, rMV As New ADODB.Recordset
' sql1 = "SELECT schgr, faelle.bhfb as bfb, faelle.Pat_id, faelle.nachname, faelle.vorname, hae.vorname as hvv, hae.nachname as hvn, hae.fax1, hausaerzte.telefax, üwnnr FROM (((faelle inner join diagnosen on faelle.pat_id = diagnosen.pat_id) left join anamnesebogen on faelle.pat_id = anamnesebogen.pat_id) left join kvaerzte.hae on concat(mid$(faelle.üwnnr,1,2)¡""/""¡mid$(faelle.üwnnr,3,5)) = hae.KVNR) left join " & DefDB & ".hausaerzte on faelle.üwnnr = hausaerzte.kvnr where quartal = """ & ZQuart(Now - Verspätung) & """ and schgr in (""24"",""23"",""21"",""00"") and tkz = 0 order by nachname, vorname, schgr"
 sql1 = "select bhfb as bfb, schgr, pat_id from faelle where schgr in (""24"",""23"",""21"",""00"") and quartal = """ & ZQuart(Now - Verspätung) & """ "
' If lies.obMySQL Then
'  sql1 = Replace(sql1, "¡", ",")
' Else
'  sql1 = Replace(Replace(sql1, "concat", ""), "¡", " & ")
' End If
 rsA.Open "select count(distinct pat_id) from (" & sql1 & ") as innen", dbcn, adOpenStatic, adLockReadOnly
 frm.Ausgabe = rsA.Fields(0) & " Fälle " & vbCrLf & altAusgabe
 altAusgabe = frm.Ausgabe
 Set rsA = Nothing
 rsA.Open sql1, dbcn, adOpenDynamic, adLockOptimistic
 Dim FZahl&, NZahl&
 Dim AltID&
 Dim anfang%, obzutr%
 Dim OkSql$
 Dim BDT As New BDTSchreib
 If Not BDT.Start(hVerz, "Leist") Then
  Exit Function
 End If
 Call BDT.BDTKopf
 Call BDT.ImportFolderHerricht
 Do While Not rsA.EOF
  obzutr = -1
  If rsA!SchGr = "00" Then ' bei eigenen nur die im DMP, da diese vermutlich noch einen anderen Hausarzt haben
    OkSql = "SELECT * FROM ((((forminhkopf left join formulare on forminhkopf.form_id = formulare.formid) left join forminhfeld on forminhfeld.foid = forminhkopf.foid) left join forminhaltfeld on forminhaltfeld.feldvw = forminhfeld.feldvw) left join forminhaltfeldinh on forminhaltfeldinh.feldinhvw = forminhfeld.feldinhvw) left join namen on forminhkopf.pat_id = namen.pat_id where form_abk like ""DMPDTYP" & "%" & """ and feld like """ & "%" & "datum"" and namen.Pat_id = " & rsA!Pat_id & ";"
    Set rMV = Nothing
    rMV.Open OkSql, dbcn, adOpenDynamic, adLockReadOnly
    If rMV.EOF Then obzutr = 0
  End If
  If obzutr Then
    Set rMV = Nothing
    rMV.Open "select * from leistungen where pat_id = " & rsA!Pat_id & " and leistung = 1601 and zeitpunkt >= " & datform(QAnf(ZQuart(Now - Verspätung))) & " and zeitpunkt < " & datform(QEnd(ZQuart(Now - Verspätung))), dbcn, adOpenDynamic, adLockReadOnly
    If Not rMV.BOF Then obzutr = 0
  End If
  If obzutr Then
   If anfang Then
    If rsA!Pat_id <> AltID Then
' in MVZ nur die schicken, die nicht schon bei mir eingeschrieben
      Dim DaT As Date
      DaT = CDate("30.6.07")
      If rsA!bfb > DaT Then
       DaT = CDate("30.6.07")
      End If
'      Call LeistungsExport1(BDT, rsA!Pat_id, "01601", DaT, CDate("18:00"))
      Call LeistungsExport1(BDT, rsA!Pat_id, "40120", DaT, CDate("18:00"))
      FZahl = FZahl + 1
    End If
    AltID = rsA!Pat_id
   End If
   anfang = True
  Else
   NZahl = NZahl + 1
  End If
'  endif
  rsA.Move 1
 Loop
 Call BDT.Schreib
 frm.Ausgabe = "In " & FZahl & " Fällen Leistungen eingetragen, in " & NZahl & " weggelassen." & vbCrLf & altAusgabe
 altAusgabe = frm.Ausgabe
 Exit Function
End Function
Function alleDMPLeiDok1(frm As Lese)
 Dim rsA As New ADODB.Recordset, rMV As New ADODB.Recordset
 
 sql1 = "SELECT faelle.bhfb as bfb, faelle.Pat_id, faelle.nachname, faelle.vorname, hae.vorname as hvv, hae.nachname as hvn, hae.fax1, hausaerzte.telefax, üwnnr FROM (((faelle inner join diagnosen on faelle.pat_id = diagnosen.pat_id) left join anamnesebogen on faelle.pat_id = anamnesebogen.pat_id) left join kvaerzte.hae on concat(mid$(faelle.üwnnr,1,2)¡""/""¡mid$(faelle.üwnnr,3,5)) = hae.KVNR) left join `" & DefDB(dbcn) & "`.hausaerzte on faelle.üwnnr = hausaerzte.kvnr where quartal = """ & ZQuart(Now - Verspätung) & """ and schgr in (""24"",""23"",""21"") and tkz = 0 and ((icd like ""E11." & "%" & """ and hae.DMPT2 <> 0) or (icd like ""E10." & "%" & """ and hae.DMPT1 <> 0)) and not üwnnr in (""6493842"",""6496648"",""6419153"",""6491291"",""6488286"",""6420029"") order by nachname, vorname, schgr;"
'' Dr. colberg / Schmidt:
' sql1 = "SELECT distinct faelle.Pat_id, faelle.nachname, faelle.vorname, hae.vorname as hvv, hae.nachname as hvn, hae.fax1, hausaerzte.telefax, üwnnr FROM (((faelle inner join diagnosen on faelle.pat_id = diagnosen.pat_id) left join anamnesebogen on faelle.pat_id = anamnesebogen.pat_id) left join kvaerzte.hae on concat(mid$(faelle.üwnnr,1,2)¡""/""¡mid$(faelle.üwnnr,3,5)) = hae.KVNR) left join " & defdb & ".hausaerzte on faelle.üwnnr = hausaerzte.kvnr where quartal = """ & ZQuart(Now - Verspätung) & """ and schgr in (""24"",""23"",""21"") and tkz = 0 and ((icd like ""E11." & "%" & """ and hae.DMPT2 <> 0) or (icd like ""E10." & "%" & """ and hae.DMPT1 <> 0)) and üwnnr in (""6419027"", ""6419421"",""6419418"") order by nachname, vorname, schgr;"
 ' Dr. Hofner 64/80054
' sql1 = "SELECT distinct faelle.Pat_id, faelle.nachname, faelle.vorname, hae.vorname as hvv, hae.nachname as hvn, hae.fax1, hausaerzte.telefax, üwnnr FROM (((faelle inner join diagnosen on faelle.pat_id = diagnosen.pat_id) left join anamnesebogen on faelle.pat_id = anamnesebogen.pat_id) left join kvaerzte.hae on concat(mid$(faelle.üwnnr,1,2)¡""/""¡mid$(faelle.üwnnr,3,5)) = hae.KVNR) left join " & defdb & ".hausaerzte on faelle.üwnnr = hausaerzte.kvnr where quartal = """ & ZQuart(Now - Verspätung) & """ and schgr in (""24"",""23"",""21"") and tkz = 0 and ((icd like ""E11." & "%" & """ and hae.DMPT2 <> 0) or (icd like ""E10." & "%" & """ and hae.DMPT1 <> 0)) and üwnnr in (""6480054"") order by nachname, vorname, schgr;"
 If lies.obMySQL Then
  sql1 = Replace(sql1, "¡", ",")
 Else
  sql1 = Replace(Replace(sql1, "concat", ""), "¡", " & ")
 End If
 rsA.Open sql1, dbcn, adOpenDynamic, adLockOptimistic
 Dim FZahl&, NZahl&
 Dim AltID&
 Dim anfang%
 Dim BDT As New BDTSchreib
 If Not BDT.Start(hVerz, "Leist") Then
  Exit Function
 End If
 Call BDT.ImportFolderHerricht
 Call BDT.BDTKopf
' Call LeistungsExport0
 Do While Not rsA.EOF
  If anfang Then
 ' If (IsNull(rsA!telefax) And IsNull(rsA!fax1)) Or (Not IsNull(rsA!telefax) And Not IsNull(rsA!fax1) And rsA!telefax <> rsA!fax1) Then
   If rsA!Pat_id <> AltID Then
' in MVZ nur die schicken, die nicht schon bei mir eingeschrieben
    Dim OkSql$
    OkSql = "SELECT * FROM ((((forminhkopf left join formulare on forminhkopf.form_id = formulare.formid) left join forminhfeld on forminhfeld.foid = forminhkopf.foid) left join forminhaltfeld on forminhaltfeld.feldvw = forminhfeld.feldvw) left join forminhaltfeldinh on forminhaltfeldinh.feldinhvw = forminhfeld.feldinhvw) left join namen on forminhkopf.pat_id = namen.pat_id where form_abk like ""DMPDTYP" & "%" & """ and feld like """ & "%" & "datum"" and namen.Pat_id = " & rsA!Pat_id & ";"
    Set rMV = Nothing
    rMV.Open OkSql, dbcn, adOpenDynamic, adLockReadOnly
    If rMV.EOF Or rsA!ÜWNNr <> "6419416" Then ' beim MVZ nicht die schon bei uns eingetragenen
     Set rMV = Nothing
     rMV.Open "select * from leistungen where leistung = 1600 and zeitpunkt >= " & datform(CDate("28.6.07")) & " and zeitpunkt < " & datform(CDate("28.6.07") + 1) & " and pat_id = " & rsA!Pat_id, dbcn, adOpenDynamic, adLockReadOnly
     If rMV.BOF Then
      Dim DaT As Date
      DaT = CDate("16.6.07")
      If rsA!bfb > DaT Then
       DaT = CDate("30.6.07")
      End If
'      Call LeistungsExport1(BDT, rsA!Pat_id&, "01600", DaT, CDate("13:00"))
      Call LeistungsExport1(BDT, rsA!Pat_id&, "40120", DaT, CDate("13:00"))
      FZahl = FZahl + 1
     Else
      NZahl = NZahl + 1
     End If
    End If
    AltID = rsA!Pat_id
   End If
  End If
'  If rsA!Pat_id = 977 Then
   anfang = True
'  endif
  rsA.Move 1
 Loop
 frm.Ausgabe = "In " & FZahl & " Fällen Leistungen eingetragen, in " & NZahl & " weggelassen." & vbCrLf & altAusgabe
 altAusgabe = frm.Ausgabe
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in alleDMPLeiDok/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' alleDMPLeiDok
Function max(d1, d2)
 If d1 > d2 Then max = d1 Else max = d2
End Function
Function min(d1, d2)
 If d1 < d2 Then min = d1 Else min = d2
End Function
Function obhierdmp%(Notiz$, Optional maxdHier As Date, Optional maxdHA As Date, Optional maxdNein As Date, Optional obdmp%)
 Dim buch$, pos%, notdat$(), testd As Date, i%
    maxdHier = 0
    maxdHA = 0
    maxdNein = 0
    If Notiz <> "" Then
     ReDim notdat(0)
     Do
      pos = InStr(Notiz, vbCrLf)
      If pos > 0 Then
       notdat(UBound(notdat)) = UCase$(Trim$(Left(Notiz, pos - 1)))
       Notiz = Mid$(Notiz, pos + 2)
       ReDim Preserve notdat(UBound(notdat) + 1)
      Else
       notdat(UBound(notdat)) = UCase$(Trim$(Notiz))
       Exit Do
      End If
     Loop
     For i = 0 To UBound(notdat)
      If InStrB(notdat(i), "DMP") <> 0 Then
       For pos = Len(notdat(i)) To 0 Step -1
        buch = Mid$(notdat(i), pos, 1)
        If buch = " " Then
         If IsDate(Mid$(notdat(i), pos + 1)) Then
          testd = CDate(Mid$(notdat(i), pos + 1))
          If InStrB(notdat(i), "HIER") <> 0 Then
           maxdHier = max(maxdHier, testd)
          ElseIf InStrB(notdat(i), " HA") <> 0 Or InStrB(notdat(i), " HÄ") <> 0 Then
           maxdHA = max(maxdHA, testd)
          ElseIf InStrB(notdat(i), "nein") <> 0 Then
           maxdNein = max(maxdNein, testd)
          End If
         End If
         Exit For
        End If
       Next pos
      End If
     Next i
    End If
 obhierdmp = (maxdHier > maxdHA) And (maxdHier > maxdNein)
 obdmp = (maxdHier > maxdNein) Or (maxdHA > maxdNein)
End Function ' obhierdmp
Function alleDMPs(frm As Lese)
 Dim docName$, Adressat$, fax1$, anfang%, runde%, dszahl&
 Dim hae As New ADODB.Recordset
 Dim rsA As New ADODB.Recordset, rMV As New ADODB.Recordset
 Dim sql$, sql1$
 Dim AltID&
 Dim antw&
 antw = MsgBox("Wollen Sie wirklich alle DMP-Daten versenden?", vbYesNo)
 If antw = vbNo Then Exit Function
 QMdbAkt = QmdB ' Imortiere.bas "u:\anamnese\quelle.mdb"
'SELECT faelle.Pat_id, faelle.nachname, faelle.vorname, hae.telefax as hfax, hausaerzte.telefax FROM (((faelle inner join diagnosen on faelle.pat_id = diagnosen.pat_id) left join anamnesebogen on faelle.pat_id = anamnesebogen.pat_id) left join kvaerzte.hae on concat(mid$(faelle.üwnnr,1,2)¡"/"¡mid$(faelle.üwnnr,3,5)) = hae.KVNR) left join " & defdb & ".hausaerzte on faelle.üwnnr = hausaerzte.kvnr where quartal = "42006" and schgr in ("24","23","21") and icd like "E11.%" and tkz = 0 and hae.DMPT2 <> 0 and not üwnnr in ("6493842","6496648","6419153","6491291","6488286","6420029","6494531") order by schgr;
'Ausnahmen: Lindenthal, Pfeuffer, Rembold, Schneider, Stolzki, Wengeler
 
 Dim Faxnr$, Infos$() ' Frau/Herrn, Vorn+Nachn, Straße, PLZ+Ort, Faxnr, S.g./Liebe, DMPTyp2, DMPTyp1
 Dim aktPatGefaxt$()
 Dim i%, j&, obdoppelt%
'  left join kvaerzte.hae on concat(mid$(faelle.üwnnr,1,2)¡""/""¡mid$(faelle.üwnnr,3,5)) = hae.KVNR) left join " & DefDB & ".hausaerzte on faelle.üwnnr = hausaerzte.kvnr) left join " & DefDB & ".listenausgabeuew as haxls on faelle.üwnnr = haxls.kvnr

 sql1 = "SELECT faelle.Pat_id, faelle.nachname, faelle.vorname, üwnnr, namen.notiz, icd FROM (((`" & DefDB(dbcn) & "`.faelle inner join diagnosen on faelle.pat_id = diagnosen.pat_id) left join anamnesebogen on faelle.pat_id = anamnesebogen.pat_id) left join namen on faelle.pat_id = namen.pat_id) where (quartal = """ & ZQuart(Now - Verspätung) & """ and schgr in (""00"",""24"") and tkz = 0 and ((icd like ""E11.%"") or (icd like ""E10.%""))) order by nachname, vorname, schgr"
'' Dr. colberg / Schmidt:
' sql1 = "SELECT distinct faelle.Pat_id, faelle.nachname, faelle.vorname, hae.vorname as hvv, hae.nachname as hvn, hae.fax1, hausaerzte.telefax, haxls.fax, üwnnr, namen.notiz FROM (((((faelle inner join diagnosen on faelle.pat_id = diagnosen.pat_id) left join anamnesebogen on faelle.pat_id = anamnesebogen.pat_id) left join kvaerzte.hae on concat(mid$(faelle.üwnnr,1,2)¡""/""¡mid$(faelle.üwnnr,3,5)) = hae.KVNR) left join " & defdb & ".hausaerzte on faelle.üwnnr = hausaerzte.kvnr) left join " & defdb & ".listenausgabeuew as haxls on faelle.üwnnr = haxls.kvnr) left join namen on faelle.pat_id = namen.pat_id where quartal = """ & ZQuart(Now - Verspätung) & """ and schgr in (""24"",""23"",""21"") and tkz = 0 and ((icd like ""E11." & "%" & """ and hae.DMPT2 <> 0) or (icd like ""E10." & "%" & """ and hae.DMPT1 <> 0)) and üwnnr in (""6419027"", ""6419421"",""6419418"") order by nachname, vorname, schgr"
 ' Dr. Hofner 64/80054
' sql1 = "SELECT distinct faelle.Pat_id, faelle.nachname, faelle.vorname, hae.vorname as hvv, hae.nachname as hvn, hae.fax1, hausaerzte.telefax, haxls.fax, üwnnr, namen.notiz FROM (((((faelle inner join diagnosen on faelle.pat_id = diagnosen.pat_id) left join anamnesebogen on faelle.pat_id = anamnesebogen.pat_id) left join kvaerzte.hae on concat(mid$(faelle.üwnnr,1,2)¡""/""¡mid$(faelle.üwnnr,3,5)) = hae.KVNR) left join " & defdb & ".hausaerzte on faelle.üwnnr = hausaerzte.kvnr) left join " & defdb & ".listenausgabeuew as haxls on faelle.üwnnr = haxls.kvnr) left join namen on faelle.pat_id = namen.pat_id where quartal = """ & ZQuart(Now - Verspätung) & """ and schgr in (""24"",""23"",""21"") and tkz = 0 and ((icd like ""E11." & "%" & """ and hae.DMPT2 <> 0) or (icd like ""E10." & "%" & """ and hae.DMPT1 <> 0)) and üwnnr in (""6480054"") order by nachname, vorname, schgr"
 If lies.obMySQL Then
  sql1 = Replace(sql1, "¡", ",")
 Else
  sql1 = Replace(Replace(sql1, "concat", ""), "¡", " & ")
 End If
 sql1 = "select * from (" & sql1 & ") as innen group by pat_id" ' where pat_id = 2193
 rsA.Open sql1, dbcn, adOpenDynamic, adLockOptimistic
' For runde = 1 To 2
 Do While Not rsA.EOF
   If rsA!Pat_id <> AltID Then
    ReDim aktPatGefaxt(0)
    If Not obhierdmp(rsA!Notiz) Then ' falls nicht bei uns unterschrieben
     Call getHausarzt(rsA!Pat_id, Infos())
     For i = 0 To UBound(Infos, 2)
      If ((rsA!ICD Like "E11*" And Infos(6, i) = "X") Or (rsA!ICD Like "E10*" And Infos(7, i) = "X")) Then
       Select Case Infos(12, i)
        Case "6419153", "6493950", "6419568" ' Schade, Schorten, Sewering
'        Case "6495818", "6491291" ' Pfeuffer, Krombholz
        Case Else
         fax1 = Infos(4, i)
         If fax1 = "08131 25169" Then fax1 = "08131 273373" ' Dr. Stolzki falsche Faxnummer
         obdoppelt = False
         For j = 0 To UBound(aktPatGefaxt) - 1
          If aktPatGefaxt(j) = fax1 Then obdoppelt = True
         Next j
         If Not obdoppelt Then
          Debug.Print "Erstelle: ", rsA!Pat_id, rsA!NachName, rsA!Vorname, fax1, rsA!ÜWNNr
'          If runde = 1 Then
'          Else
'           If rsA!Pat_id = 1974 Then
           anfang = True
'           End If
           If anfang Then
            Call ausgeb(rsA!Pat_id & ": " & rsA!NachName & ", " & rsA!Vorname & ", " & fax1 & ", " & rsA!ÜWNNr)
            DoEvents
            Adressat = Infos(1, i)
            docName = do_DMPAusgebStandAlone(rsA!Pat_id, fax1, Adressat)
            If False Then
             Call FaxSend(docName, Adressat, fax1)
            End If
           End If ' 1 = 0
           aktPatGefaxt(UBound(aktPatGefaxt)) = fax1
           ReDim Preserve aktPatGefaxt(UBound(aktPatGefaxt) + 1)
'         Exit For ' Alternative zu obdoppelt
'          End If
          frm.Bytes = dszahl
          dszahl = dszahl + 1
         End If ' runde = 1
       End Select
      End If
     Next i
    End If
   End If
   AltID = rsA!Pat_id
  rsA.Move 1
 Loop
 rsA.MoveFirst
 frm.GesBytes = dszahl
 dszahl = 0
' Next runde
 lies.Ausgabe = "Fertig mit alle DMPs!" & vbCrLf & altAusgabe
 altAusgabe = lies.Ausgabe
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in alleDMPs/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' allDMPs
Function ausgeb(AGStr$)
            Open AutoBriefProtok For Append As #334 ' p:\zufaxen.txt = AutoBriefProtok
            Print #334, AGStr
            Close #334
 Lese.Ausgabe = AGStr & vbCrLf & altAusgabe
 altAusgabe = Lese.Ausgabe
End Function
Sub dodoFollowUp(frm As Lese)
 Dim Infos$()
 Dim sql$, sql1$
 Dim rP As New ADODB.Recordset
 Dim rB As New ADODB.Recordset
 Dim Zp As Date, Quartal$
 On Error GoTo fehler
 Zp = Now - 30
 Quartal = ZQuart(Zp)
' ab = InputBox("Ab welchem Pat. anfangen (0 = von vorne)?", "Rückfrage", 0)
 Call rP.Open("select distinct f.pat_id as pat_id from faelle f left join anamnesebogen a on f.pat_id = a.pat_id where not schgr in ('41','42','43')  and tkz = 0 order by f.pat_id desc;", dbcn, adOpenDynamic, adLockReadOnly)
 Do While Not rP.EOF
  Call getHausarzt(rP!Pat_id, Infos())
  If Infos(1, 0) <> "" And Not Infos(1, 0) Like "*Schade" Then
   Set rB = Nothing
   sql = "SELECT * From briefe WHERE pat_id = " & rP!Pat_id & " and ((name Like '%Brief an %Dr%' Or name Like '%Arztbrief%' Or name Like 'Brief an HA%' Or name Like 'Brief an HAe%' or name like 'Brief an %') And name Not Like '%Entwurf%') order by zeitpunkt desc"
   Call rB.Open(sql, dbcn, adOpenDynamic, adLockReadOnly)
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
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doRestlicheBriefe/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub
Sub doUngeschriebeneBriefe(frm As Lese)
 Dim Infos$()
 Dim sql$, sql1$
 Dim rP As New ADODB.Recordset
 Dim rB As New ADODB.Recordset
 Dim Zp As Date, Quartal$
 On Error GoTo fehler
 Zp = Now - 30
 Quartal = ZQuart(Zp)
' ab = InputBox("Ab welchem Pat. anfangen (0 = von vorne)?", "Rückfrage", 0)
 Call rP.Open("select distinct f.pat_id as pat_id from faelle f left join anamnesebogen a on f.pat_id = a.pat_id where not schgr in ('41','42','43')  and tkz = 0  order by f.pat_id desc;", dbcn, adOpenDynamic, adLockReadOnly)
 Do While Not rP.EOF
  Call getHausarzt(rP!Pat_id, Infos())
  If Infos(1, 0) <> "" And Not Infos(1, 0) Like "*Schade" Then
   Set rB = Nothing
   sql = "SELECT * From briefe WHERE pat_id = " & rP!Pat_id & " and ((name Like '%Brief an %Dr%' Or name Like '%Arztbrief%' Or name Like 'Brief an HA%' Or name Like 'Brief an HAe%' or name like 'Brief an %') And name Not Like '%Entwurf%')"
   Call rB.Open(sql, dbcn, adOpenDynamic, adLockReadOnly)
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
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doRestlicheBriefe/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
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
 Call rP.Open("select distinct pat_id from faelle where quartal = '" & Quartal & "' and pat_id >= " & ab & " order by pat_id ", dbcn, adOpenDynamic, adLockReadOnly)
 Do While Not rP.EOF
  Set rB = Nothing
  sql = "SELECT * From briefe WHERE pat_id = " & rP!Pat_id & " and zeitpunkt > " & datform(QAnf(Quartal)) & " and ((name Like '%Brief an %Dr%' Or name Like '%Arztbrief%' Or name Like 'Brief an HA%' Or name Like 'Brief an HAe%') And name Not Like '%Entwurf%')"
  Call rB.Open(sql, dbcn, adOpenDynamic, adLockReadOnly)
  If rB.BOF Then
   Call tubriefStandalone(rP!Pat_id, True)
  End If
  rP.Move 1
 Loop
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doRestlicheBriefe/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
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
Dim DateiName$
Dim strFaxTiff As FAXCOMLib.FaxTiff
Dim strJobID&
On Error GoTo fehler
 DateiName = "u:\test1.txt"
 Err.Clear
 FaxServer.Connect ("anmeld2")

Set FaxDoc = FaxServer.CreateDocument(DateiName)
   
    FaxDoc.BillingCode = "Rechnungsnummer 381"
    FaxDoc.CoverpageName = ""
    FaxDoc.CoverpageNote = "Insulinanaloga"
    FaxDoc.CoverpageSubject = "Insulinanaloga"
    FaxDoc.DiscountSend = 0
    FaxDoc.DisplayName = "G.Schade"
    FaxDoc.EmailAddress = "diabetologie@dachau-mail.de"
    FaxDoc.FaxNumber = RecNum
    FaxDoc.RecipientAddress = ""
    FaxDoc.RecipientCity = ""
    FaxDoc.RecipientCompany = "Praxis"
    FaxDoc.RecipientCountry = "D"
    FaxDoc.RecipientDepartment = ""
    FaxDoc.RecipientHomePhone = ""
    FaxDoc.RecipientName = RecName
    FaxDoc.RecipientOffice = ""
    FaxDoc.RecipientOfficePhone = ""
    FaxDoc.RecipientState = "Bayern"
    FaxDoc.RecipientTitle = ""
    FaxDoc.RecipientZip = ""
    FaxDoc.SendCoverpage = 0
    FaxDoc.SenderAddress = "Mittermayerstraße 13"
    FaxDoc.SenderCompany = "Diabetologische Schwerpunktpraxis"
    FaxDoc.SenderDepartment = "Schreibbüro"
    FaxDoc.SenderFax = "08131 616381"
    FaxDoc.SenderHomePhone = "616380"
    FaxDoc.SenderName = "Gerald Schade"
    FaxDoc.SenderOffice = "Praxis"
    FaxDoc.SenderOfficePhone = "616380"
    FaxDoc.SenderTitle = ""
    FaxDoc.ServerCoverpage = 1
    FaxDoc.Filename = docName
    FaxDoc.DisplayName = docName
    strJobID = FaxDoc.Send
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
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FaxSend/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub ' FaxSend

Public Function do_DMPAusgebStandAlone(Pat_id&, Optional fax1$, Optional Adressat$)
 Dim dc, VorString$, docName$, DT As dmptyp
 Dim mR1, mR2, mR3
' Dim raNa As DAO.Recordset
 Dim rAna As New ADODB.Recordset
 Dim sverz$
 If fax1$ = "" Then
  sverz = "p:\"
 Else
  sverz = "p:\zufaxen\"
 End If
 On Error GoTo fehler
' dtbInit
' Set raNa = TabÖff("Anamnesebogen", "Pat_id")
 rAna.Open "select * from anamnesebogen where pat_id = " & Pat_id, dbcn, adOpenStatic, adLockReadOnly
' raNa.Seek "=", Pat_id
 VorString = ""
 nzw = vbCr
' If rAna!DMPhier <> CDate(0) And rAna!HAimDMP = "Hausarzt im DMP" Then
'  VorString = "Bei mir fand am " + format$(rAna!DMPhier, "dd/mm/yyyy") + " die Einschreibung ins DMP statt. Diese sollte jedoch auf Sie als Hausarzt übertragen werden." + nzw
' End If
 GetWord
 With Wapp
  .Options.SmartCutPaste = False
  On Error Resume Next
  .Options.SmartParaSelection = False
  On Error GoTo fehler
  Call .Documents.Add("u:\vorlagen\DMP-Vorlage.dot")
  Set dc = .ActiveDocument
  dc.Range.InsertAfter VorString & DMPString$(Pat_id, DT)
       Set mR1 = dc.content
       With mR1.Find
         .clearformatting
         .Text = "DMP-Informationen vom"
         .Replacement.Text = ""
         .wrap = wdFindContinue
         .Format = False
         .Execute
       End With
       If mR1.Find.found Then
        Set mR2 = dc.Range(mR1.Start, mR1.Start)
        mR2.Find.Text = ":"
        mR2.Find.Execute
        If mR2.Find.found Then
         Set mR3 = dc.Range(mR1.Start, mR2.End)
'         mR3.Font.Bold = True
        On Error Resume Next
        Dim para
        Set para = mR3.Paragraphs.First.Range
        Do While Err.Number = 0
         Set para = para.Paragraphs.First.Next.Range
         With dc.Range(para.Start, para.Start + InStr(para.Text, ":")).Font
          .Italic = True
          .Bold = True
         End With
        Loop
        End If
       End If
  On Error GoTo fehler
  
  dc.sections(2).Range.ParagraphFormat.TabStops.ClearAll
  dc.sections(2).Range.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(5.5), Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderDots
  dc.sections(2).Range.ParagraphFormat.FirstLineIndent = -35
  dc.sections(2).Range.ParagraphFormat.LeftIndent = 35
  docName = sverz + rAna!NachName & " " & rAna!Vorname & ", PID " & rAna!Pat_id & ", DMP-Daten vom " + Format$(Now, "DD/MM/YY hh.mm.ss") & IIf(Adressat <> "" And Not IsNull(Adressat), " für " & Adressat, "") & IIf(fax1 = "", "", " an Fax " & fax1) & ".doc"
  dc.SaveAs Filename:=docName
  dc.Close
'  .Visible = True
  
'  .Application.WindowState = wdWindowStateMaximize
'  .Activate
 End With
 do_DMPAusgebStandAlone = docName
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in do_DMPAusgebStandAlone/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' do_DMPAusgebStandalone

Public Function ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung
Dim j As String * 4, q As String * 1
On Error GoTo fehler
j = Year(Datum)
Select Case Datum
 Case Is < CDate("1.4." + j): q = "1"
 Case Is < CDate("1.7." + j): q = "2"
 Case Is < CDate("1.10." + j): q = "3"
 Case Else: q = "4"
End Select
ZQuart = q + j
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZQuart/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung

Function GesNam$(ByVal rs) ' As DAO.Recordset) ' s.a. GesName(
 On Error GoTo fehler
   GesNam = rs!NachName & ", " & IIf(IsNull(rs!Titel), "", rs!Titel) + IIf(IsNull(rs!Titel) Or rs!Titel = "", "", ", ") + rs!Vorname
   On Error Resume Next
   GesNam = rs!NVorsatz + IIf(rs!NVorsatz = "", "", " ") + GesNam
   On Error GoTo fehler
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GesNam/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
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
Function daseit$(ra As ADODB.Recordset)
 Dim rs As New ADODB.Recordset
 If IsNull(ra("Diabetes seit")) Then
  daseit = "?"
 ElseIf LCase(ra("Diabetes seit")) = "bu" Then
  rs.Open "select fanf from faelle where fid = (select min(fid) from faelle where pat_id = " & ra!Pat_id & ")", dbcn, adOpenKeyset, adLockReadOnly
'  daseit = format$(rs!Fanf, "mm\/yy")
   daseit = Format$(rs!Fanf, "yyyy")
 Else
  daseit = ra("Diabetes seit")
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DASeit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DAseit
Function DSeit(ranam As ADODB.Recordset)
 Dim rsDS As New ADODB.Recordset
 On Error GoTo fehler
    If IsNull(ranam("Diabetes seit")) Then
      DSeit = "?"
    ElseIf LCase(ranam("Diabetes seit")) = "bu" Then
'      DSeit = format$(rAnam!Vorgestellt, "mm\/yy")
'      DSeit = format$(Dtb.OpenRecordset("select fanf from faelle where fid = (select min(fid) from faelle where pat_id = " + CStr(rAnam!Pat_id) + ")")!Fanf, "mm\/yy")
      rsDS.Open "select fanf from faelle where fid = (select min(fid) from faelle where pat_id = " + CStr(ranam!Pat_id) + ")", dbcn, adOpenStatic, adLockReadOnly
      If Not rsDS.EOF Then
       DSeit = Format$(rsDS!Fanf, "mm\/yy")
      End If
    Else: DSeit = CStr(ranam("Diabetes seit"))
    End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DSeit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DSeit(rAnam)

Function TabPr(s1, s2)
'Dim nzl$
'nzl = vbcr + vblf ' vbcrlf kannte ich da noch nicht
 DmPStr = DmPStr & vbCrLf & s1
 If Len(s1) < 24 Then
  DmPStr = DmPStr & Space$(24 - Len(s1))
 End If
 DmPStr = DmPStr & vbTab & IIf(IsNull(s2), "", s2) ' "........................"
End Function ' TabPr(S1, S2)

Function obPosi(S) As Boolean
 If IsNull(S) Then
  obPosi = False
 ElseIf VarType(S) = vbBoolean Then
  obPosi = S
 Else
  Select Case tfeld(S)
   Case "", "n", "nie", "-", 0, "nein", "normal", "keine | keine", "- | -", "|", "n | n"
    obPosi = False
   Case Else
    obPosi = True
  End Select
 End If
End Function ' obPosi(S) As Boolean
' wird in TherArt() aufgerufen, wo anhand des Diabetestyps und von Formularen Pumpentherapien ermittelt werden, sodann hiermit die übrigen Therapieformen
Function TherAuskunft(Pat_id, obanf%, Optional insz%, Optional vordat As Date, Optional obIns As Boolean, Optional obAnal As Boolean, Optional obGlib As AntidiabMedType, Optional obmetf As AntidiabMedType, Optional obGlucI As AntidiabMedType, Optional obSHGlin As AntidiabMedType, Optional obGlit As AntidiabMedType, Optional obSonstAD As AntidiabMedType, Optional obHMG As Boolean, Optional obAntihyp As Boolean, Optional obACEH As Boolean, Optional obBetabl As Boolean, Optional obThro As Boolean, Optional obAntikoag As Boolean, Optional Qmax$, Optional obRezIns As Boolean, Optional DT$) As TherapieArt
' Dim rMA As DAO.Recordset
 Dim obNIns 'ob Normalinsulin
 Dim raMa As New ADODB.Recordset
 Dim AktMz%, medi, lzpos%
 Dim MPNr&(), MP0&, MPe&, MPz&, i&, MPNrl& ' erste und letzte, laufende MedPlan-Nummer des Quartals
 Dim Thfakt As TherapieArt
 On Error GoTo fehler
 If LenB(dbcn) = 0 Then Call Lese.ProgStart
' QuartalMax: Falls befüllt, so wird die höchstwertige Therapie des in der Variablen angegebenen Quartals ermittelt
' Set rMa = TabÖff("MedArten", "Medikament")
'Medikamente:
 MPz = 0
 ReDim MPNr(MPz)
 If Qmax <> "" Then
  MP0 = MedPlanNr(Pat_id, obAkt:=True, vordat:=QAnf(Qmax), NurNr:=True)
  MPe = MedPlanNr(Pat_id, obAkt:=True, vordat:=QEnd(Qmax) + 1, NurNr:=True)
  Set raMa = Nothing
  Call raMa.Open("Select distinct mpnr from medplan where pat_id = " & Pat_id & " and mpnr >= " & MP0 & " and mpnr <= " & MPe & " order by mpnr", dbcn, adOpenStatic, adLockReadOnly)
  Do While Not raMa.EOF
   MPz = MPz + 1
   ReDim Preserve MPNr(MPz)
   MPNr(MPz) = raMa!MPNr
   raMa.Move 1
  Loop
 Else
  MPz = 1
  ReDim MPNr(MPz)
  If obanf Then
   Call MedPlanAusAna(Pat_id)
  ElseIf obanf = 1 Then
   MPNr(0) = MedPlanNr(Pat_id, -1)
  Else
   MPNr(0) = MedPlanNr(Pat_id, Not obanf, vordat)
  End If
 End If
 TherAuskunft = offen
 For i = 1 To MPz 'MPNrl = MPNrA To MPNrE
  If Qmax <> "" Then
   Call MedPlanNr(Pat_id, True, , , nr:=MPNr(i)) ' dann die übrigen Variablen hier jeweils neu befüllen
  End If
  obGlib = 0: obmetf = 0: obGlucI = 0: obSHGlin = 0: obGlit = 0: obSonstAD = 0
  obIns = 0: obAnal = 0: obHMG = 0: obAntihyp = 0: obACEH = 0: obBetabl = 0: obThro = 0
  insz = 0
  obNIns = 0
  For AktMz = 0 To MedZahl
   Dim DosH%
   DosH = DosHäuf(AktMz)
'  If DosH > 0 Then ' wenn auch bei Dosierung was drinsteht
   medi = Trim$(LCase(Med(AktMz)))
   If medi <> "" And medi <> "-" Then
    lzpos = InStr(medi, " ")
    If lzpos > 0 Then medi = Left(medi, lzpos - 1)
'    rMA.Seek "=", medi
    Set raMa = Nothing
    raMa.Open "select -hmg as j_hmg, -hypt as j_hypt, -aceh as j_aceh, -betabl as j_betabl, -thro as j_thro, -antikoag as j_antikoag, -glib as j_glib, -metf as j_metf, -gluci as j_gluci, -shglin as j_shglin, -glit as j_glit, -sonstad as j_sonstad, -ins as j_ins, -AnAl as j_anal, m.* from medarten m where medikament = '" & medi & "'", dbcn, adOpenDynamic, adLockReadOnly
    If raMa.BOF Then
     If DosH > 0 Then
'      MsgBox "Medikament: " + Medi + " noch nicht in Tabelle MedArten erfaßt"
      Debug.Print "Medikament: " + medi + " noch nicht in Tabelle MedArten erfaßt"
      dbcn.Execute ("insert into medarten(langname,medikament,hinzugefügt,pat_id) values('" & Med(AktMz) & "','" & UCase$(medi) & "'," & datform(Now) & "," & Pat_id & ")")
     End If
    Else
     If Not IsNull(raMa!j_antikoag) Then If raMa!j_antikoag <> 0 Then obAntikoag = True ' Bei Marcumar steht meist keine Dosierung drin
     If DosH > 0 Then
      If Not IsNull(raMa!j_glib) Then If raMa!j_glib <> 0 Then obGlib = True
      If Not IsNull(raMa!j_Metf) Then If raMa!j_Metf <> 0 Then obmetf = True
      If Not IsNull(raMa!j_GlucI) Then If raMa!j_GlucI <> 0 Then obGlucI = True
      If Not IsNull(raMa!j_SHGlin) Then If raMa!j_SHGlin <> 0 Then obSHGlin = True
      If Not IsNull(raMa!j_Glit) Then If raMa!j_Glit <> 0 Then obGlit = True
      If Not IsNull(raMa!j_SonstAD) Then If raMa!j_SonstAD <> 0 Then obSonstAD = True
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
      If Not IsNull(raMa!j_betabl) Then If raMa!j_betabl <> 0 Then obBetabl = True
      If Not IsNull(raMa!j_aceh) Then If raMa!j_aceh <> 0 Then obACEH = True
      If Not IsNull(raMa!j_thro) Then If raMa!j_thro <> 0 Then obThro = True
     End If ' dosh > 0
    End If ' rma.nomatch
   End If
'  End If 'doshäuf(aktmz)
  Next AktMz ' aktmz = 0 to mednr
  Thfakt = offen
  If obNIns Then
    Thfakt = ict
  Else
    Select Case insz
     Case 0
      If obRezIns Then
       Thfakt = ict
      Else
       If obGlib Or obmetf Or obGlucI Or obSHGlin Or obGlit Or obSonstAD Then
        Thfakt = oad
       Else
        If DT = "1" Then
         Thfakt = ct
        Else
         Thfakt = diät
        End If
       End If
      End If
     Case 1
      If obGlib Or obGlucI Or obSHGlin Or obGlit Or obmetf Or obSonstAD Then
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
   If Thfakt > TherAuskunft Then TherAuskunft = Thfakt
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
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in TherAuskunft/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' TherAuskunft(pat_id, obanf%, insz, Optional obIns, Optional obAnal, Optional obGlib, Optional obMetf, Optional obGlucI, Optional obSHGlin, Optional obGlit, Optional obSonstAD, Optional obHMG, Optional obAntihyp, Optional obThro)
Function diI(icd_str, Optional pid, Optional abDat As Date, Optional auchZ As Boolean, Optional Weiteres$) As Boolean
 Dim raDT As New ADODB.Recordset, raFa As New ADODB.Recordset
 Dim diIsql$
 On Error GoTo fehler
 If Not IsMissing(pid) Then Pat_id = pid
 If Pat_id = 0 Then
  MsgBox "Pat_id = 0 in diI(.. => Programmfehler"
  Stop
  Exit Function
 End If
 diIsql = "select " & IIf(lies.obMySQL, "", "top 1 ") & "* from diagnosen where pat_id = " & Pat_id & " and icd like '" & icd_str & "%" & "' and diagsicherheit in ('G','V'" & IIf(auchZ, ",'Z'", "") & ") " & IIf(abDat <> 0, " and diagdatum >= " & datform(abDat), "") & IIf(LenB(Weiteres) <> 0, " and " & Weiteres, "")
 Dim lddat As Date
 raFa.Open "select max(bhfb) as lddat from faelle where pat_id = " & Pat_id, dbcn, adOpenStatic, adLockReadOnly
 If Not raFa.BOF Then
  lddat = raFa!lddat
  diIsql = diIsql & " and (obdauer or diagdatum >= " & datform(lddat) & ")"
 End If
 If lies.obMySQL Then diIsql = diIsql & " limit 1"
 Call raDT.Open(diIsql, dbcn, adOpenStatic, adLockReadOnly)
 diI = Not raDT.BOF
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in diI/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function 'diI(icd_str, PID) As Boolean
Function diT(diag_str, Optional pid) As Boolean
 Dim raDT As New ADODB.Recordset
 Dim raFa As New ADODB.Recordset
 Dim diTsql$
 If Not IsMissing(pid) Then Pat_id = pid
 If Pat_id = 0 Then
  MsgBox "Pat_id = 0 in diT(.. => Programmfehler"
  Stop
  Exit Function
 End If
 diTsql = "select " & IIf(lies.obMySQL, "", "top 1 ") & "* from diagnosen where pat_id = " & Pat_id & " and diagtext like '" & "%" & diag_str & "%" & "' and diagsicherheit in (""G"",""V"")"
 Dim lddat As Date
 raFa.Open "select max(bhfb) as lddat from faelle where pat_id = " & Pat_id, dbcn, adOpenStatic, adLockReadOnly
 If Not raFa.BOF Then
  lddat = raFa!lddat
  diTsql = diTsql & " and (obdauer or diagdatum >= " & datform(lddat) & ")"
 End If
 If lies.obMySQL Then diTsql = diTsql & " limit 1"
 Call raDT.Open(diTsql, dbcn, adOpenStatic, adLockReadOnly)
 diT = Not raDT.BOF
End Function 'diT(icd_str, PID) As Boolean

'Function Seite$()
'  Select Case raDT!DiagSeite
'  Case "R": Seite = "rechts"
'  Case "L": Seite = "links"
'  Case "B": Seite = "bds."
' End Select
'End Function ' seite$()

Function PulsParse$(Feld, Optional repath As Boolean, Optional lipath As Boolean)
     Dim Add$, re$, li$
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
     If InStrB(Add, "/") = 0 Then Add = Replace(Add, ",", "/")
     Add = Replace(Replace(Replace(Replace(Replace(Replace(Add, "biph.", "biphasisch "), "biphas.", "biphasisch "), "hyperä.", "hyperämisch "), "hyperä", "hyperämisch "), "mono", "monophasisch"), "pst", "poststenotisch ")
     If Right$(Add, 2) = "bi" Then Add = Add + "phasisch"
     If InStrB(Add, ",") = 0 And InStrB(Add, "/") = 0 And InStrB(Add, "bds") = 0 And InStrB(Add, "beid") = 0 Then
      Add = "bds. " + Add
     End If
     PulsParse = Add
     re = Add
     li = Add
     If InStrB(Add, "/") <> 0 Then
      re = Left(Add, InStr(Add, "/") - 1)
      li = Mid$(Add, InStr(Add, "/") + 1)
     End If
     If InStrB(Add, "|") <> 0 Then
      re = Left(Add, InStr(Add, "|") - 1)
      li = Mid$(Add, InStr(Add, "|") + 1)
     End If
     If InStrB(re, "-") <> 0 Or InStrB(re, "?") <> 0 Then repath = True
     If InStrB(li, "-") <> 0 Or InStrB(li, "?") <> 0 Then lipath = True
End Function ' PulsParse

Function tfeld$(S)
If IsNull(S) Then
 tfeld = ""
Else
 tfeld = S
End If
End Function ' TFeld$(S)

Function obMonPath(Feld, rep%, lip%)
Dim TestFeld$, Tre$, Tli$
rep = 1
lip = 1
SensSplit Feld, Tre, Tli
If InStrB(Tre, "/5") <> 0 Then
 Tre = Replace(Tre, "/5", "")
 Tli = Replace(Tli, "/5", "")
 If Tre = "" Or Tre = "~" Or InStrB(Tre, "5") <> 0 Then rep = 0
 If Tli = "" Or Tli = "~" Or InStrB(Tli, "5") <> 0 Then lip = 0
ElseIf InStrB(Tre, "/3") <> 0 Then
 Tre = Replace(Tre, "/3", "")
 Tli = Replace(Tli, "/3", "")
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
Tre = Replace(Tre, "/5", "")
Tli = Replace(Tli, "/5", "")
If Tre = "" Or Tre = "~" Or InStrB(Tre, "4") <> 0 Or InStrB(Tre, "5") <> 0 Then rep = 0
If Tli = "" Or Tli = "~" Or InStrB(Tli, "4") <> 0 Or InStrB(Tli, "5") <> 0 Then lip = 0
End Function ' obKWPath(Feld, rep%, lip%)
Function obVibPath(Feld, rep%, lip%, Optional obStreng%)
Dim TestFeld$, Tre$, Tli$
rep = 1
lip = 1
SensSplit Feld, Tre, Tli
Tre = Replace(Tre, "/8", "")
Tli = Replace(Tli, "/8", "")
If (Tre = "" Or Tre = "~" Or (obStreng = 0 And InStrB(Tre, "4") <> 0) Or InStrB(Tre, "5") <> 0 Or InStrB(Tre, "6") <> 0 Or InStrB(Tre, "7") <> 0 Or InStrB(Tre, "8") <> 0) And InStrB(Tre, ">8") = 0 Then rep = 0
If (Tli = "" Or Tli = "~" Or (obStreng = 0 And InStrB(Tli, "4") <> 0) Or InStrB(Tli, "5") <> 0 Or InStrB(Tli, "6") <> 0 Or InStrB(Tli, "7") <> 0 Or InStrB(Tli, "8") <> 0) And InStrB(Tli, ">8") = 0 Then lip = 0
End Function ' obVibPath(Feld, rep%, lip%)


Function SensSplit(SensStrg, re$, li$)
If IsNull(SensStrg) Then SensStrg = ""
If InStrB(SensStrg, "|") <> 0 Then
      re = LTrim$(Left(CStr(SensStrg), InStr(CStr(SensStrg), "|") - 1))
      li = LTrim$(Mid$(CStr(SensStrg), InStr(CStr(SensStrg), "|") + 1))
ElseIf InStrB(SensStrg, ",") <> 0 Then
      re = Left(CStr(SensStrg), InStr(CStr(SensStrg), ",") - 1)
      li = Mid$(CStr(SensStrg), InStr(CStr(SensStrg), ",") + 1)
Else
      re = SensStrg
      li = SensStrg
End If
re = Trim$(re)
li = Trim$(li)
End Function ' SensSplit(SensStrg, re$, li$)

 Public Function GetPrRR$(rs As ADODB.Recordset, RRsyst%, RRdiast%)
 '  Dim rrr As DAO.Recordset
'  Set rrr = TabÖff("RR", "Auswahl")
'  rrr.Seek "=", CStr(rs!Pat_id)
  Dim rarr As New ADODB.Recordset
  On Error GoTo fehler
  GetPrRR = ""
  rarr.Open "select * from rr where pat_id = " & rs!Pat_id & " order by zeitpunkt desc", dbcn, adOpenDynamic, adLockReadOnly
  With rs
  If Not rarr.EOF Then
   GetPrRR = rarr!RR & IIf(InStrB(rarr!RR, "mm Hg") = 0 And InStrB(rarr!RR, "mmHg") = 0, " mm Hg", "") + " (" + Format$(rarr!Zeitpunkt, "DD.MM.YY") + ")"
  ElseIf tfeld(!RR) <> "" Then
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
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GetPrRR/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' GetPrRR
#If False Then
Function SchulzBest%(Pat_id&, zp1$, zpl$, Optional abDat)
  Dim sqls$, lapp As New ADODB.Recordset
  On Error GoTo fehler
  sqls$ = "SELECT * FROM eintraege where pat_id = " & CStr(Pat_id) + " and art = ""schul"" "
  If Not IsMissing(abDat) Then If Now > CDate("15.10.05") Then sqls = sqls + "and Zeitpunkt >= " & datform(abDat) & " "
  sqls = sqls + "order by zeitpunkt"
  lapp.Open sqls, dbcn, adOpenDynamic, adLockReadOnly
  If Not lapp.BOF Then
   lapp.MoveLast
   zpl = Format$(lapp!Zeitpunkt, "dd/mm/yy")
   lapp.MoveFirst
   zp1 = Format$(lapp!Zeitpunkt, "dd/mm/yy")
  End If
  lapp.Close
  lapp.Open "select count(*) as zl from (" & sqls & ")" & IIf(lies.obMySQL, " as innen", ""), dbcn, adOpenStatic, adLockReadOnly
  SchulzBest = lapp!zl
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SchutzBest/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'SchulzBest
#End If

Function rrEmpf$(RRsyst%, RRdiast%, Pat_id)
  Dim gwSyst%, gwDiast%
  If Not diI("I10", Pat_id) Then
   If RRsyst > 149 Or RRdiast > 100 Then
    If MsgBox("Achtung: Fehlt Diagnose Hypertonie bei Pat_id = " & CStr(Pat_id) & "?", vbYesNo) = vbYes Then
     MsgBox "Bitte in Turbomed eintragen"
    End If
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

Function DatInStr(TestStr$, Optional Jahr) As Date
 Dim buch As String * 1, SL%, aktb%, p1%, pL%, puZ%, ZwiStr$
 SL = Len(TestStr)
 p1 = 0
 pL = 0
 puZ = 0
 For aktb = 1 To SL
  buch = Mid$(TestStr, aktb, 1)
  If p1 = 0 Then
   If InStrB("0123456789", buch) <> 0 And buch <> "" Then p1 = aktb
  End If
  If p1 > 0 Then
   If buch = "." Then puZ = puZ + 1
   If InStrB("0123456789" + IIf(puZ > 2, "", "."), buch) = 0 Then
    pL = aktb - p1
    Exit For
   End If
  End If
 Next
 DatInStr = CDate(0)
 If p1 > 0 And pL > 0 Then
  ZwiStr = Mid$(TestStr, p1, pL)
  If Not IsDate(ZwiStr) Then
   ZwiStr = ZwiStr + CStr(Jahr)
  End If
  If IsDate(ZwiStr) Then
   DatInStr = CDate(ZwiStr)
  End If
 End If
End Function ' DatInStr(TestStr$, Optional jahr) As Date

Function MedPlanAusAna(Pat_id)
 Dim rsNa As New ADODB.Recordset, i%, j%, inkl%, k%, dosstr$, dosier$()
 
' Set rsna = TabÖff("Anamnesebogen", "Pat_id")
 rsNa.Open "select * from anamnesebogen where pat_id = " & Pat_id, dbcn, adOpenStatic, adLockReadOnly
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
      dosstr = ""
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
      dosier(k) = Replace(dosier(k), "€", "-")
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
Function MedPlanNr&(Pat_id, obAkt, Optional vordat As Date, Optional NurNr%, Optional nr&)
 Dim raFIM As New ADODB.Recordset
 On Error GoTo fehler
' sql1 = "select " & IIf(obAkt, "max", "min") & "(mpnr) from medplan where pat_id = " + CStr(Pat_id)
' If vordat <> CDate(0) Then sql1 = sql1 + " and zeitpunkt < " & datform(vordat + 1)
' sql1 = sql1 + " order by zeitpunkt" + IIf(obAkt, " desc", "") + ", Medikament, feldnr"
 'Set rFIM = Dtb.OpenRecordset(sql1, dbOpenDynaset)
 If nr <> 0 Then
  MedPlanNr = nr
 Else
  sql1 = "select " & IIf(obAkt, "max", "min") & "(mpnr) as mpmax from medplan where pat_id = " + CStr(Pat_id)
  If vordat <> CDate(0) Then sql1 = sql1 + " and zeitpunkt < " & datform(vordat + 1)
  raFIM.Open sql1, dbcn, adOpenDynamic, adLockReadOnly
' If Not IsNull(raFIM.Fields(0)) Then
  If Not raFIM.BOF And Not IsNull(raFIM!mpmax) Then
   MedPlanNr = raFIM!mpmax
   If NurNr Then Exit Function
  Else
   Exit Function
  End If
 End If
 sql1 = "select * from medplan where mpnr = " & MedPlanNr
 Set raFIM = Nothing
 raFIM.Open sql1, dbcn, adOpenDynamic, adLockReadOnly
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
     Dos(0, MedNr) = IIf(IsNull(raFIM!mo), "", raFIM!mo)
     Dos(1, MedNr) = IIf(IsNull(raFIM!mi), "", raFIM!mi)
     Dos(2, MedNr) = IIf(IsNull(raFIM!nm), "", raFIM!nm)
     Dos(3, MedNr) = IIf(IsNull(raFIM!ab), "", raFIM!ab)
     Dos(4, MedNr) = IIf(IsNull(raFIM!Zn), "", raFIM!Zn)
    End If
    raFIM.Move 1
   Loop
  End If ' not bof
' End If ' not isnull
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description + vbCrLf + "Fehlerposition: " + CStr(FPos), vbAbortRetryIgnore, "Aufgefangener Fehler in MedPlanNr/" + App.path)
 Case vbAbort: Call MsgBox("Höre auf"): End
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
     If T1 = "" Then GoTo nix
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
        If InStrB("0123456789", Mid$(T1a, i, 1)) <> 0 And Mid$(T1a, i, 1) <> "" Then Exit For
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
     erg = ""
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
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dodoRRParse/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dodoRRParse
Function WieTabak(Text$, Optional Pat_id&) As ZigSt
  Dim raZig As New ADODB.Recordset
  WieTabak = doTabakSt(Text)
  If Pat_id <> 0 Then
   Call raZig.Open("select * from diagnosen where pat_id = " & Pat_id & " and icd like '" & "F17" & "%" & "' and diagsicherheit in (""G"",""V"",""Z"")", dbcn, adOpenDynamic, adLockReadOnly)
   If Not raZig.EOF Then
    If InStrB(raZig!DiagText, "Z.n") <> 0 Or raZig!DiagSicherheit = "Z" Then
     If WieTabak <> Früher And WieTabak <> VorLangem Then WieTabak = Früher
    Else
     WieTabak = Aktuell
    End If
   End If
  End If
End Function ' WieTabak

 Function doTabakSt(Text) As ZigSt  ' TabakStatus "parsen"
  Dim KommaStelle%, FrüherStelle%, bisStelle%, MinusStelle%, LJStelle% ' 0 = keiner, 1 = früher, 2 = aktuell, 3 = vor mehr als 15 Jahren
  Dim jahrk$, obdatum%, Datu As Date, Sp$()
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
    Sp = Split(Text, "Frühergeraucht?")
    If Trim$(Sp(0)) = ",wieviel?" Or obNein(Left(Sp(0), 1)) Then
     doTabakSt = Nie
    Else
     If rSnpos > 0 Then
      If obNein(rSn) Then doTabakSt = Früher Else If (obNein(Sp(0)) Or InStrB(rSn, "j") <> 0) Then doTabakSt = Aktuell
     Else
      If LCase(Left(Text, 1)) = "j" Then
       If InStrB(Text, "is") <> 0 Then
        doTabakSt = Früher
       Else
        doTabakSt = Aktuell
       End If
      ElseIf obNein(Sp(0)) Then
       doTabakSt = Aktuell
      ElseIf UBound(Sp) < 1 Then
       doTabakSt = Nie
      ElseIf obNein(Sp(1)) Then
       doTabakSt = Nie
      Else
       doTabakSt = Früher
      End If
     End If
    End If
   ElseIf Left(rSn, 1) = "j" Then 'obNein(rSn) = 0 Then 4.4.07 Pat. 1716
    doTabakSt = Aktuell
   ElseIf Left(Text, 1) = "n" Or Text = "Frühergeraucht? n" Or Text = "Frühergeraucht? -" Or Left(Text, 1) = "," Then
    doTabakSt = Nie
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
     doTabakSt = Früher
    ElseIf (FrüherStelle > 0 And KommaStelle > 0 And FrüherStelle > KommaStelle) Then
     doTabakSt = Aktuell
    Else
     If InStrB(Text, "ja") <> 0 Or InStr(Text, "j") = 1 Then
      doTabakSt = Aktuell
     Else
      If Text = "n" Or Text = "-" Or InStrB(Text, "nein") <> 0 Or InStrB(Text, "kein") <> 0 Or Text = "0" Then
       doTabakSt = Nie
      Else
       If (InStrB(Text, "LJ") <> 0 And MinusStelle > 0) Or (MinusStelle = 1 And obdatum) Then
'        If Not obdatum Or (obdatum And datu > Now + 15 * 365) Then
' muß noch getestet werden
         doTabakSt = Früher
       Else
        doTabakSt = Aktuell
       End If
      End If
     End If
    End If
   End If
  End If ' not isnull(Text
  End If
   If doTabakSt = Früher And InStrB(Text, "biswann:") <> 0 Then
    bisStelle = InStr(Text, "biswann:") + Len("biswann:")
    jahrk = LTrim$(Trim$(Mid$(Text, bisStelle + 1)))
    obdatum = -1
   End If
  If doTabakSt = 1 And obdatum Then
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
    doTabakSt = VorLangem
   End If
  End If
    
  Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doTabakst/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
 End Function ' doTabakst

Function ZSU$(Quelle$)
Dim i%, ZS$, bq$, bz$
 On Error GoTo fehler
ZS = ""
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
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZSU/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ZSU
