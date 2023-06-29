Attribute VB_Name = "WerteVerarbeiten"
Option Explicit
'Public Const artspez$ = "(art IN (""notiz"",""ni"",""telef"",""gs"",""rz"",""ga"",""ag"",""hz"",""ts"",""cr"",""ep"",""ke"")"
Dim DmPStr$
'Dim rDT AS DAO.Recordset
Dim SStr$
Dim Pat_id&
'Enum DFSNiveau
'   stNichts%
'   St0%
'   St1%
'   St2%
'   Ampzeh%
'   St3%
'   AmpUS%
'   St4%
'   AmpOS%
'  END Enum

Sub werteAnzeig()
Dim Abfr As New ADODB.Recordset
Dim i&, runde%
Dim bisher()
Dim bCt%, aktbCt%
Dim Dat0$, Dat1$
Dat0$ = aVerz & "Allewerte.txt"
Dat1$ = aVerz & "Verschiedenewerte.txt"
Dim DaT(2)
DaT(0) = Dat0
DaT(1) = Dat1
For runde = 0 To 1
' SET Abfr = Dtb.OpenRecordset(Anmnb, dbOpenDynaset)
' Abfr.Open "SELECT * FROM `anakt`", DBCn, adOpenDynamic, adLockReadOnly
 myFrag Abfr, "SELECT * FROM `anakt`"
 If Not Abfr.BOF Then
  Open DaT(runde) For Output Access Write Lock Write As #1
  For i = 1 To Abfr.Fields.COUNT - 1
   Print #1, Abfr.Fields(i).name
   Abfr.MoveFirst
   bCt = 0
   ReDim bisher(0)
   While Not Abfr.EOF
    If Not IsNull(Abfr.Fields(i).Value) Then
     If Not Abfr.Fields(i).Value = "" Then
      If runde = 1 Then
       For aktbCt = 0 To bCt - 1
        If bisher(aktbCt) = Abfr.Fields(i).Value Then GoTo gibts
       Next
       bisher(bCt) = Abfr.Fields(i).Value
       bCt = bCt + 1
       ReDim Preserve bisher(bCt)
      End If
      Print #1, "  " + CStr(Abfr.Fields(i).Value)
gibts:
     End If
    End If
    Abfr.MoveNext
   Wend
  Next i
 End If
 Abfr.Close
 Close #1
 zeigan CStr(DaT(runde)), vbNormalFocus
Next runde
End Sub ' werteAnzeig()
#If False Then
Sub wAusgeb()
werteAusGeb True
End Sub ' wAusgeb()
Sub werteAusGeb(Optional obStumm As Boolean = False)
'Dim ID$
'Dim tbl
'' DoCmd.Save acForm, Anmnb geht leider nicht
'On Error Resume Next
'ID = forms!Anamnesebogen!Pat_ID
'On Error GoTo fehler
'If ID = "" THEN
' For Each tbl In dtb.TableDefs
'  DoCmd.Close acTable, tbl.Name, acSaveYes
' Next
' DoCmd.OpenForm dtb.Containers(2).Documents(0).Name 'Anamnesebogen
' DoCmd.maximize
' ID = forms!Anamnesebogen!Pat_ID
'END IF
Dim rs As New ADODB.Recordset
Dim DT As DMPClass
Call Formmerk
Call FormAufruf
Call GibwerteAus(Not obStumm)
Call DiagnosenAusgeb(AktID, Not obStumm)
'Set rs = TabÖff("Anamnesebogen", "Pat_id")
'rs.Seek "=", AktID
myFrag rs, "SELECT Diabetestyp FROM `anakt` WHERE pat_id = " & CStr(AktID)
If rs.BOF Or (Not rs.BOF And (rs!Diabetestyp = "2" Or rs!Diabetestyp = "s")) Then
 Call DMPAusgeb0(DT, CStr(AktID), Not obStumm)
End If
Call behDatAusgeb(AktID, Not obStumm)
Call LaborAusgeb(AktID)
Call FormAufruf
End Sub ' werteAusGeb
#End If
#If False Then
Sub DiagnosenAusgeb(Optional id&, Optional obAnzeig As Boolean = True)
Const ErgebDatei$ = aVerz + "Diagnosen.txt"
'Dim EStr$
Dim Tbl
' DoCmd.Save acForm, Anmnb geht leider nicht
If id = 0 Then
 On Error Resume Next
 Call AnbogVar(True)
 id = Forms(Anmnbi)(ABPat_ID)
 On Error GoTo fehler
 If id = 0 Then
  For Each Tbl In Dtb.TableDefs
   DoCmd.Close acTable, Tbl.name, acSaveYes
  Next
  DoCmd.OpenForm Dtb.Containers(2).documents(0).name
  DoCmd.Maximize
  Call AnbogVar(True)
  id = Forms(Anmnbi)(ABPat_ID)
 End If
End If
Open ErgebDatei For Output Access Write Lock Write As #1
'EStr = replace$(Forms!Anamnesebogen!Diagnosen, " `", vbtab + "`")
'EStr = LEFT(EStr, Len(EStr) - 2)
Print #1, AusgDiag(id)
Close #1
If obAnzeig Then zeigan ErgebDatei, vbNormalFocus
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DiagnosenAusgeb/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' Diagnosenausgeb
#End If

Function AusgDiag$(Pat_id&, Optional ohneNotwend%)
 Dim rsNa As ADODB.Recordset
 Dim Spl$(), j%
' SET rsNa = TabÖff("Anamnesebogen", "Pat_ID")
 Set rsNa = Nothing
 myFrag rsNa, "SELECT Diagnosen FROM `anamnesebogen` WHERE pat_id = " & Pat_id
 If Not rsNa.EOF Then
  Spl = Split(rsNa!Diagnosen, vbVerticalTab)
  For j = 0 To UBound(Spl)
   If ohneNotwend Then If InStrB(Spl(j), "Notw") > 0 Then Spl(j) = vNS
   Spl(j) = REPLACE$(Spl(j), " `", vbTab + "`")
   If Not obkNeph Or InStrB(Spl(j), "N08.3") = 0 Then ' Nephropathie ggf. streichen
    If Spl(j) <> "" Then AusgDiag = AusgDiag + Spl(j) + vbVerticalTab
   End If
  Next j
'  AusgDiag = replace$(rsAnam!Diagnosen, " `", vbtab + "`")
  If Len(AusgDiag) >= 1 Then AusgDiag = Left$(AusgDiag, Len(AusgDiag) - 1)
 End If
 rsNa.Close
End Function ' AusDiag

#If False Then
Sub behDatAusgeb(Optional id&, Optional obAnzeig As Boolean = True)
Dim lbehD As Date, Tbl As tabledef
'Dim rNo As New ADODb.Recordset
Const ErgebDatei$ = aVerz + "behDat.txt"
On Error GoTo fehler
If id = 0 Then
 On Error Resume Next
 Call AnbogVar
 id = Forms(Anmnbi)(ABPat_ID)
 On Error GoTo fehler
 If id = 0 Then
  For Each Tbl In Dtb.TableDefs
   DoCmd.Close acTable, Tbl.name, acSaveYes
  Next
  DoCmd.OpenForm Dtb.Containers(2).documents(0).name
  DoCmd.Maximize
  id = Forms(Anmnbi)(ABPat_ID)
 End If
End If
Pat_id = id
Open ErgebDatei For Output Access Write Lock Write As #2
lbehD = lebe(Pat_id)
Print #2, CStr(Forms(Anmnbi).adoRS!Vorgestellt) + IIf(lBehDat <> CDate(0), ", zuletzt am " + Format$(lbehD, "dd/mm/yy") + ",", "")
Close #2
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in behDatAusgeb/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' behDatAusgeb
#End If

' aufgerufen in lebe, behDauerStr, UKPDS
Function erbe(ByVal Pat_id&) As Date
'Static rFa AS DAO.Recordset
'Set rFa = TabÖff("faelle", "ErstF")
Dim rsAnam As New ADODB.Recordset
On Error GoTo fehler
'Call rsAnam.Open("SELECT aufndat FROM `namen` WHERE pat_id = " & Pat_id, DBCn, adOpenDynamic, adLockOptimistic)
myFrag rsAnam, "SELECT aufndat FROM `namen` WHERE pat_id = " & Pat_id
If rsAnam.EOF Then Exit Function
erbe = rsAnam!AufnDat
If erbe > #7/1/2004# Then Exit Function

Dim rFaName$
Dim raFa As New ADODB.Recordset
'rFaName = raFa.Name
'On Error GoTo fehler
'If rFaName = "" THEN SET rFa = TabÖff("faelle", "Auswahl")
myFrag raFa, "SELECT * FROM `faelle` WHERE pat_id = " & Pat_id & " AND fanf > " & DatFor_k("1.7.2004") & " ORDER BY fanf"

'rFa.Seek "=", Pat_id
' erbe = IIf(rFa!SchGr = "90", rFa!bhFb, IIf(rFa!ausgst = CDate(0), rFa!lVorl, rFa!ausgst))

If Not raFa.EOF Then
' IF rFa!Pat_id <> Pat_id THEN Exit Do
 If raFa!Fanf < erbe Or erbe = 0 Then erbe = raFa!Fanf
Else
 Call doFAnfFuell(raFa)
End If
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in erbe/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' erbe(byVal Pat_id&) As Date

' in lebetest
Function lebe(ByVal Pat_id&) As Date ' Letzte Behandlung
Dim lbehD As Date, rLb As New ADODB.Recordset
On Error GoTo fehler
lebe = erbe(Pat_id)
'lbehD = Dtb.OpenRecordset("SELECT MAX(zeitpunkt) AS zp FROM `" + QmdbAkt + "`.`medplan` WHERE pat_id = " + CStr(Pat_id), dbOpenDynaset)!zp
Set rLb = Nothing
myFrag rLb, "SELECT MAX(zeitpunkt) AS zp FROM `medplan` WHERE pat_id = " + CStr(Pat_id)
If rLb.EOF Then
 Exit Function
Else
 If Not IsNull(rLb!Zp) Then
  lbehD = rLb!Zp
 End If
End If
If lbehD > lebe Then lebe = lbehD
'lbehD = Dtb.OpenRecordset("SELECT MAX(zeitpunkt) AS zp FROM `" + QmdbAkt + "`.`eintraege` WHERE (" & artspezG & ") OR LEFT(art,2)=""rr"" OR LEFT(art,2)=""bz"") AND pat_id = " + CStr(Pat_id), dbOpenDynaset)!zp
Set rLb = Nothing
myFrag rLb, "SELECT MAX(zeitpunkt) AS zp FROM `eintraege` WHERE " + "(art IN (" & artspezG & ") OR LEFT(art,2)=""rr"" OR LEFT(art,2)=""bz"") AND pat_id = " + CStr(Pat_id)
If Not rLb.BOF Then
 If Not IsNull(rLb!Zp) Then
  lbehD = rLb!Zp
 End If
End If
If lbehD > lebe Then lebe = lbehD
lBehDat = lbehD
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in lebe/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' letzt

Function behDauerStr$(Pat_id, lbeh As Date) ' nur in tubriefStandalone
 Dim D1 As Date, D2 As Date
 On Error GoTo fehler
 D1 = DateValue(erbe(Pat_id))
 D2 = lbeh ' DateValue(lebe(Pat_id))
 If D1 = D2 Then
  behDauerStr = "am " + Format$(D1, "d.m.YY")
 Else
  behDauerStr = "vom " + Format$(D1, "d.m.YY") + " bis zum " + Format$(D2, "d.m.YY") + " mehrmals"
 End If
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in behDauerStr/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' behDauerStr$()

Sub LaborAusgeb(Optional id&)
 Dim ErgebDatei$
 ErgebDatei$ = aVerz + "Labor.txt"
 On Error GoTo fehler
 If id = 0 Then
  On Error Resume Next
  Call AnbogVar(True)
  id = Forms(Anmnbi)(ABPat_ID)
  On Error GoTo fehler
 End If
 Close #15
 Open ErgebDatei For Output Access Write Lock Write As #15
 Print #15, LaborString$(id)
 Close #15
 zeigan ErgebDatei, vbNormalFocus
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborAusgeb/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' LaborAusgeb
 
Function LaborString$(Pat_id&)
 Dim raLw As New ADODB.Recordset, raDat As New ADODB.Recordset, ls$, raLU As New ADODB.Recordset
' zeilenzahl bestimmen
 Dim ZZ&, rz$, gschl$, Vgl$, altGruppe%, Nb$ ' Normbereich
 Dim u!, o! ' oberer und unterer Grenzwert numerisch
 Dim uNG$, oNG$ ' obere und untere Normgrenze in zeichen
 Dim pKz$ ' Pathologisch-Kennzeichen
 Dim abküz%, gruppez%
 'dtbInit
 Call Lese.ProgStart
' sql = "SELECT Pat_ID, zeitPunkt, FertigStGrad, AbKü, LangText,wert, Einheit, Kommentar,"""" AS Nb FROM (" & laborAbfr & " WHERE pat_id = " & Pat_id & ") AS labor UNION SELECT Pat_ID, Eingang AS zeitpunkt, befArt AS FertigStGrad, Abkü, langname AS Langtext, wert, Einheit, Kommentar, Nb " + _
  "FROM `laborxus` LEFT JOIN laborxwert ON laborxus.RefNr=laborxwert.RefNr LEFT JOIN laborxpnb nb ON laborxwert.nbid = nb.id" + _
  "WHERE pat_id = " + CStr(Pat_id) + " AND NOT EXISTS (SELECT * FROM `laborneu` WHERE pat_id = " + CStr(Pat_id) + " AND abkü = laborxwert.Abkü AND wert = laborxwert.wert AND zeitpunkt > laborxus.Eingang-3 AND zeitpunkt < laborxus.Eingang+6)"
' sql = "SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_id & " UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_id & ") i GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb ORDER BY zeitpunkt DESC"
' sql = "SELECT lab.* FROM (SELECT @patid:=" & Pat_id & " nix) nul, geslab lab;"
'Set rLU = Dtb.OpenRecordset(sql)
' SET raLU = Nothing
' raLU.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 Set raLU = hollabor(Pat_id)
 
 gschl = vNS
 altGruppe = 0
 Set raLw = Nothing
 myFrag raLw, "SELECT 0 FROM `namen` WHERE pat_id = " & Pat_id
' SET rLw = TabÖff("Namen", "Pat_id")
' ralw.Seek "=", Pat_id
 If Not raLw.EOF Then
  gschl = raLw!geschlecht
 End If
'#Const problematisch = True
#If problematisch Then
 Set raLw = Nothing
 sql1 = "SELECT COUNT(0) FROM (SELECT DISTINCT abkü FROM geslab i) i"
' SET rLw = Dtb.OpenRecordset(sql1)
 myFrag raLw, sql1
' SET rLw = Dtb.OpenRecordset("SELECT COUNT(0) FROM (SELECT DISTINCT abkü  FROM `" + QmdbAkt + "`.laborUNION WHERE pat_id = " + CStr(Pat_id) + ")")
 ZZ = raLw.Fields(0)
' sql1 = "SELECT COUNT(0) FROM (SELECT DISTINCT gruppe FROM (SELECT * FROM (" & sql & ") sql LEFT JOIN `laborparameter` ON sql.abkü = `laborparameter`.`abkü` WHERE pat_id = " & CStr(Pat_id) & "))"
 sql1 = "SELECT COUNT(DISTINCT gruppe) FROM geslab"
' SET rLw = Dtb.OpenRecordset(sql1)
' SET rLw = Dtb.OpenRecordset("SELECT COUNT(0) FROM (SELECT DISTINCT gruppe FROM (SELECT * FROM `" + QmdbAkt + "`.LaborUNION LEFT JOIN `" + QmdbAkt + "`.`laborparameter` ON LaborUnion.abkü = `laborparameter`.`abkü` WHERE pat_id = " + CStr(Pat_id) + "))")
 Set raLw = Nothing
 myFrag raLw, sql1
 ZZ = ZZ + raLw.Fields(0)
 raLw.Close
' sql1 = "SELECT COUNT(0) FROM (SELECT distinct(DATE(zeitpunkt))  FROM `" & QMdbAkt & "`.labor LEFT JOIN `laborparameter` ON Labor.abkü = `laborparameter`.abkü WHERE pat_id = " & CStr(Pat_id) & ") i"
' sql1 = "SELECT COUNT(0) FROM (SELECT distinct(DATE(zeitpunkt)) FROM `laborneu` l LEFT JOIN `laborparameter` p ON l.abkü = p.abkü WHERE pat_id = " & CStr(Pat_id) & ")i"
 sql1 = "SELECT DISTINCT zeitpunkt FROM geslab"
' IF lies.obMySQL THEN sql1 = replace$(sql1, "datevalue(", "date(")
' SET raDat = Dtb.OpenRecordset(sql1)
' SET raDat = Dtb.OpenRecordset("SELECT COUNT(0) FROM (SELECT distinct(cdate(int(zeitpunkt)))  FROM `" + QmdbAkt + "`.labor LEFT JOIN `" + QmdbAkt + "`.`laborparameter` ON Labor.abkü = `laborparameter`.abkü WHERE pat_id = " + CStr(Pat_id) + ")")
 Set raDat = Nothing
 myFrag raDat, sql1
 rz = raDat.Fields(0)
 raDat.Close
 'statt datserial
' sql1 = "SELECT * FROM (SELECT distinct(DATE(zeitpunkt)) AS Datum FROM (" & sql & ") i LEFT JOIN `laborparameter` p ON i.abkü = p.abkü WHERE pat_id = " & CStr(Pat_id) & ") i ORDER BY datum"
 sql1 = "SELECT zp FROM (SELECT DATE(zeitpunkt) zp FROM geslab)i GROUP BY zp ORDER BY zp"
' IF lies.obMySQL THEN sql1 = replace$(sql1, "datevalue(", "date(")
' SET raDat = Dtb.OpenRecordset(sql1)
' SET raDat = Dtb.OpenRecordset("SELECT * FROM (SELECT distinct(cdate(int(zeitpunkt))) AS Datum FROM `" + QmdbAkt + "`.LaborUNION LEFT JOIN `" + QmdbAkt + "`.`laborparameter` ON LaborUnion.abkü = `laborparameter`.abkü WHERE pat_id = " + CStr(Pat_id) + ") ORDER BY datum")
 Set raDat = Nothing
' raDat.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
 myFrag raDat, sql1
 ls = CStr(rz) & vbCrLf
#Else
 ZZ = myEFrag("CALL geslabkatz(" & Pat_id & ",'abkü')")!Zahl + _
      myEFrag("CALL geslabkatz(" & Pat_id & ",'gruppe')")!Zahl
 rz = myEFrag("CALL geslabkatz(" & Pat_id & ",'zeitpunkt')")!Zahl
 myFrag raDat, "CALL geslabkatg(" & Pat_id & ",'zeitpunkt')"
#End If
 If raDat.BOF Then Exit Function
 
' raDat.moveFirst
 Do While Not raDat.EOF
  ls = ls + Format$(raDat.Fields(0), "dd.mm.yyyy") & vbCrLf
  raDat.Move 1
 Loop
 
 ls = ls + CStr(ZZ) & vbCrLf
 Set raLw = Nothing
#If problematisch Then
' sql1 = "SELECT * FROM (SELECT *,DATE(zeitpunkt) AS Datum  FROM (" & sql & ") i LEFT JOIN `laborparameter` p ON i.abkü = p.abkü WHERE pat_id = " & CStr(Pat_id) & ") ORDER BY gruppe,reihe,datum"
' sql1 = "SELECT * FROM (SELECT i.*,p.gruppe,p.reihe,p.unm,p.onm,p.unw,p.onw, DATE(zeitpunkt) AS Datum  FROM (SELECT * FROM (SELECT * FROM `labor2a` WHERE pat_id = " & Pat_id & " UNION SELECT * FROM `labor1a` WHERE pat_id = " & Pat_id & ") i GROUP BY pat_id,zeitpunkt,abkü,wert,einheit,nb ORDER BY zeitpunkt DESC) i LEFT JOIN `laborparameter` p ON i.abkü = p.abkü WHERE pat_id = " & Pat_id & ") i ORDER BY gruppe,reihe,datum"
 sql1 = "SELECT * FROM (SELECT p.*, DATE(zeitpunkt) AS Datum  FROM geslab p) i ORDER BY gruppe,reihe,datum"
' IF lies.obMySQL THEN sql1 = replace$(sql1, "datevalue(", "date(")
' SET rLw = Dtb.OpenRecordset(sql1)
'' SET rLw = Dtb.OpenRecordset("SELECT * FROM (SELECT *,DATE(zeitpunkt) AS Datum  FROM `" + QmdbAkt + "`.LaborUNION LEFT JOIN `" + QmdbAkt + "`.`laborparameter` ON LaborUnion.abkü = laborparameter.abkü WHERE pat_id = " + CStr(Pat_id) + ") ORDER BY gruppe,reihe,datum")
' raLw.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
  myFrag raLw, sql1
#Else
 myFrag raLw, "CALL geslabdp(" & Pat_id & ",'GROUP BY zeitpunkt DESC,abkü,einheit ORDER BY gruppe,reihe,zeitpunkt')"
#End If
' ralw.moveFirst
 Do While Not raLw.EOF
   If Not IsNull(raLw!gruppe) Then
    If Val(raLw!gruppe) <> altGruppe Then
     ls = ls & vbCrLf
    End If
   End If
   ls = ls + raLw.Fields("LangText") & vbCrLf
   ls = ls + IIf(IsNull(raLw.Fields("einheit")), vNS, raLw.Fields("Einheit")) & vbCrLf
#If problematisch Then
   ls = ls + Format$(raLw!Datum, "dd.mm.yyyy") & vbCrLf
#Else
   ls = ls + Format$(raLw!Zeitpunkt, "dd.mm.yyyy") & vbCrLf
#End If
   ls = ls + IIf(IsNull(raLw!Wert), vNS, raLw!Wert) & vbCrLf
   u = -99
   o = -99
   pKz = "n" ' normal
   uNG = vNS
   oNG = vNS
#If False Then
   If gschl = "w" Then
    If (Not IsNull(raLw!unw) And raLw!unw <> "") Then
     uNG = REPLACE$(raLw!unw, "1:", "")
    End If
    If (Not IsNull(raLw!onw) And raLw!onw <> "") Then
     oNG = REPLACE$(raLw!onw, "1:", "")
    End If
   End If
   If uNG = "" And oNG = "" Then
    If Not IsNull(raLw!unm) Then uNG = REPLACE$(raLw!unm, "1:", "")
    If Not IsNull(raLw!onm) Then oNG = REPLACE$(raLw!onm, "1:", "")
   End If
#End If
   If Not IsNull(raLw!uNG) Then uNG = REPLACE$(raLw!uNG, "1:", "")
   If Not IsNull(raLw!oNG) Then oNG = REPLACE$(raLw!oNG, "1:", "")
   If uNG <> "" Then u = CDbl(uNG)
   If oNG <> "" Then o = CDbl(oNG)
'   IF IsNumeric(ralw!wert) OR ralw!wert LIKE "1:*" THEN
'    IF gschl = "w" THEN
'     IF Not ralw!unw = "" THEN u = CDbl(replace$(ralw!unw, "1:", ""))
'     IF Not ralw!onw = "" THEN o = CDbl(replace$(ralw!onw, "1:", ""))
'    END IF
'    IF gschl = "m" OR u = -99 THEN
'     IF Not ralw!unm = "" THEN u = CDbl(replace$(ralw!unm, "1:", ""))
'    END IF
'    IF gschl = "m" OR o = -99 THEN
'     IF Not ralw!onm = "" THEN o = CDbl(replace$(ralw!onm, "1:", ""))
'    END IF
   If IsNumeric(raLw!Wert) Or raLw!Wert Like "1:*" Then
    If u <> -99 And CDbl(REPLACE$(REPLACE$(raLw!Wert, "1:", ""), ".", ",")) < u Then pKz = "p"
    If o <> -99 And CDbl(REPLACE$(REPLACE$(raLw!Wert, "1:", ""), ".", ",")) > o Then pKz = "p"
   End If
'   END IF
   ls = ls + pKz & vbCrLf
#If False Then
   Nb = vNS
   If uNG <> "" Or oNG <> "" Then
    If gschl = "w" And (raLw!unw <> "" Or raLw!onw <> "") Then
     Nb = IIf(IsNull(raLw!unw), vNS, raLw!unw) + "-" + IIf(IsNull(raLw!onw), vNS, raLw!onw)
    End If
   End If
   If Nb = "" And (raLw!unm <> "" Or raLw!onm <> "") Then
    Nb = IIf(IsNull(raLw!unm), vNS, raLw!unm) + "-" + IIf(IsNull(raLw!onm), vNS, raLw!onm)
   End If
#End If
   If Nb = "" And (raLw!uNG <> "" Or raLw!oNG <> "") Then
    Nb = IIf(IsNull(raLw!uNG), vNS, raLw!uNG) + "-" + IIf(IsNull(raLw!oNG), vNS, raLw!oNG)
   End If
   ls = ls + Nb & vbCrLf
   If IsNull(raLw!gruppe) Then
    altGruppe = 0
   Else
    altGruppe = Val(raLw!gruppe)
   End If
   raLw.Move 1
 Loop
 
 raDat.Close
 raLw.Close
 LaborString = ls
' Open "c:\le.txt" For Output AS #2
' Print #2, LS
' Close #2
 End Function ' LaborString
 ' noch in werteAusGeb und doMachDMPBogen
Public Sub DMPAusgeb0(aktDC As DMPClass, Optional Pat_id$, Optional obAnzeig As Boolean = True, Optional obnurDMPString = False, Optional DokuDat As Date)
 'DMPAusgeb0 =
 Dim DMPStrhier$
 Dim errn%
 DMPStrhier = DMPString$(CLng(Pat_id), aktDC, , , DokuDat, Not obnurDMPString)
 If obnurDMPString Then Exit Sub
 Dim ErgebDatei$
 ErgebDatei$ = pVerz & "dmp\" & Pat_id & " DmP.txt"
 On Error Resume Next
 Kill ErgebDatei
 Err.Clear
 Open ErgebDatei For Output Access Write Lock Write As #1
 errn = Err.Number
 On Error GoTo fehler
 If errn <> 0 Then
  Open REPLACE$(ErgebDatei, "p:", "\\linux1\daten\Patientendokumente") For Output Access Write Lock Write As #1
 End If
 Print #1, DMPStrhier
 Close #1
 If obAnzeig Then zeigan ErgebDatei, vbNormalFocus
 Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description & " bei der Ausgabe in: " & ErgebDatei, vbAbortRetryIgnore, "Aufgefangener Fehler in DMPAusgeb0/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub      ' DMPAusgeb0

#If False Then
Function TherapieArtEinzelnFestlegen(Pat_id&, Optional rsAna As ADODB.Recordset) ' in TherapieArtenFestlegen und alleSpeichern
Dim nTher$, rAF&, rAnPatID&, Anzeige$, Fanf As Date
On Error Resume Next
Fanf = rsAna!Fanf
On Error GoTo fehler
Call syscmd(acSysCmdSetStatus, "Therapiearten festlegen für Patient " & Pat_id & ", Fallanfang: " & Format(Fanf, "d.m.yy"))
On Error Resume Next
rAnPatID = rsAna!Pat_id
On Error GoTo fehler
If rAnPatID <> Pat_id Then
 Set rsAna = New ADODB.Recordset
' rsAna.Open "SELECT pat_id,nachname, vorname, diabetestyp,-insulinpumpe AS j_insulinpumpe,ther1,therakt FROM `anamnesebogen` WHERE pat_id = " & Pat_id, DbCn, adOpenStatic, adLockOptimistic
 Lese.ProgStart
 myFrag rsAna, "SELECT pat_id,nachname, vorname, diabetestyp,insulinpumpe, ther1,therakt FROM `anamnesebogen` WHERE pat_id = " & Pat_id, adOpenStatic, DBCn, adLockOptimistic
 If rsAna.BOF Then Exit Function
End If ' rAnPatID
On Error GoTo fehler
If Not rsAna.EOF Then
 nTher = TherUmw(therart_erm(Pat_id, -1, rsAna))
 If nTher <> rsAna!Ther1 Or IsNull(rsAna!Ther1) Then
  Anzeige = Pat_id & ": Ther1: " & rsAna!Ther1 & " -> " & nTher
  Call syscmd(acSysCmdSetStatus, Anzeige)
  Call myEFrag("UPDATE `anamnesebogen` SET ther1 = '" & nTher & "' WHERE pat_id = " & Pat_id, rAF)
  If rAF <> 1 Then
   Anzeige = "   Fehler beim Update von Ther1 bei " & Pat_id & ": rAF = " & rAF
   Call syscmd(acSysCmdSetStatus, Anzeige)
  End If
 End If
 nTher = TherUmw(therart_erm(Pat_id, 0, rsAna))
 If nTher <> rsAna!TherAkt Or IsNull(rsAna!TherAkt) Then
'  Anzeige = Pat_id & " " & rsAna!NachName & " " & rsAna!VorName & ": TherAkt: " & rsAna!TherAkt & " -> " & nTher
  Anzeige = Pat_id & " " & ": TherAkt: " & rsAna!TherAkt & " -> " & nTher
  Call syscmd(acSysCmdSetStatus, Anzeige)
  Call myEFrag("UPDATE `anamnesebogen` SET therakt = '" & nTher & "' WHERE pat_id = " & Pat_id, rAF)
  If rAF <> 1 Then
   Anzeige = "   Fehler beim Update von TherAkt bei " & Pat_id & ": rAF = " & rAF
   Call syscmd(acSysCmdSetStatus, Anzeige)
  End If
 End If
' IF rsana!Pat_id = 0 THEN
'  rsana.CancelUpdate ' kommt auf mir noch unbekanntem weg zustande
'  Err.Raise 999, , "Fehler in TherapieArtEinzelnFestlegen: pat_id 0"
' Else
'  rsana.Update
' END IF
End If
syscmd acSysCmdClearStatus
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in TherArt/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' TherapieArtEinzelnFestlegen(Pat_id&)
#End If

#If False Then
Function testTherArt$(Optional Pat_id&, Optional obanf%)
 Dim rsNa As New ADODB.Recordset
 Call Lese.ProgStart
 If Pat_id = 0 Then Pat_id = 748
 testTherArt = TherUmw(therart_erm(Pat_id, obanf, rsNa, CDate("17.11.04")))
 Call Lese.ProgEnde
End Function ' testtherart$()
#End If

#If False Then
'Public FUNCTION therart_erm(Pat_id&, Optional obanf%, Optional rsAna As Adodb.Recordset, Optional VorDat As Date, Optional Qmax$) AS TherapieArt
'' in Therapiearteneinzelnfestlegen, Epikrise, DMPString, testTherArt
'Dim insz%, obIns AS Boolean, obAnal AS Boolean, obGlib AS AntidiabMedType, obmetf AS AntidiabMedType, obGlucI AS AntidiabMedType, obSHGlin AS AntidiabMedType, obGlit AS AntidiabMedType, obdpp4 AS AntidiabMedType, obglp1 AS AntidiabMedType, obsglt2 AS AntidiabMedType, obSonstAD AS AntidiabMedType
'Dim rAnPatID&, neuladen%
'On Error Resume Next
'therart_erm = offen
'If rsAna Is Nothing THEN
' neuladen = True
'Else
' IF rsAna!Pat_id <> Pat_id THEN neuladen = True
'END IF
'If neuladen THEN
'  SET rsAna = New ADODB.Recordset
' Call Lese.ProgStart
' rsAna.Open "SELECT pat_id,diabetestyp,insulinpumpe, ther1,therakt FROM `anamnesebogen` WHERE pat_id = " & Pat_id, DBCn, adOpenKeyset, adLockReadOnly
' IF rsAna.BOF THEN Exit Function
'END IF
'rAnPatID = rsAna!Pat_id
'On Error GoTo fehler
'         Dim DT$
'         IF ISNULL(rsAna!Diabetestyp) THEN
'          DT = 2
'         Else
'          DT = rsAna!Diabetestyp
'         END IF
' SELECT CASE LCase$(DT)
'  Case 1, 2, "g", "s", "p"
'       Dim rmed As New ADODB.Recordset, rMA As New ADODB.Recordset, obRezPumpe AS Boolean, obRezIns AS Boolean
''      dim Kurzmed$
'       Dim rlar As New ADODB.Recordset
'       IF Not obanf THEN ' aus den Rezepten über Insulin, Analoga und Pumpe Rückschlüsse ziehen, falls nicht obanf
'        Dim lzp As Date
'        lzp = Now
'        SET rlar = Nothing
'        myFrag rlar, "SELECT MAX(bhfb) AS mbhfb FROM `faelle` WHERE pat_id = " & Pat_id
'        IF Not rlar.EOF THEN
'         lzp = IIf(ISNULL(rlar!mbhfb), Now(), rlar!mbhfb)
'        END IF
'        SET rlar = Nothing '
'        'txtmedKey'
'        myFrag rlar, forminhalt & " WHERE form_abk IN ('lar','plar') AND feld IN ('medikament','txtMedKey') AND " & _
'        "(feldinh LIKE '%rapid d link%' OR feldinh LIKE '%rap d li%' OR feldinh LIKE '%rapid-d li%' OR feldinh LIKE '%tenderl%' OR feldinh LIKE '%flexl%' OR feldinh LIKE '%check spirit%' OR feldinh LIKE '%insight%' OR feldinh LIKE '%chek spirit%' OR feldinh LIKE '%pumpenträg%' OR feldinh LIKE '%kunststoffampu%' OR feldinh LIKE '%spritzampull%' OR feldinh LIKE '%batteriefachdeckel%' OR feldinh LIKE '%h-tron%' OR feldinh LIKE '%d-tron%' OR feldinh LIKE '%paradigm%' OR feldinh LIKE '%csii%' OR feldinh LIKE '%linpumpe%' OR feldinh LIKE '%omnipod%' OR Or feldinh LIKE "ypso pump" feldinh LIKE '%minimed%' OR feldinh LIKE '%640g%' OR feldinh LIKE '%carelink%' OR feldinh LIKE '%mio %' OR feldinh LIKE '%quick%set%' OR feldinh LIKE '%silhouette%' OR feldinh LIKE '%sure-t%' OR feldinh LIKE '%sure t%' OR feldinh LIKE '%paradigm%' OR feldinh LIKE '% veo%' OR feldinh LIKE '%animas%' OR feldinh LIKE '%car%idge%') " & _
'        "and zeitpunkt > " & DatFor_k(lzp - 640) & " AND pat_id = " & Pat_id
'
'        IF rlar.BOF THEN
'         obRezPumpe = False
'        Else
'         obRezPumpe = True
'        END IF
'        obRezIns = False
'        IF Not obRezPumpe THEN
'         Dim Ausschluß$, krzmd$
'         Ausschluß = " AND NOT medikament RLIKE 'senso|\ ttr|clix|loe|aviva'"
'         krzmd = "IF(INSTR(MID(medikament,4),' ')=0,medikament,LEFT(medikament,2+INSTR(MID(medikament,4),' '))) kurzmed, "
'         sql1 = "SELECT DISTINCT " & krzmd & "medikament FROM `rezepteintraege`  WHERE pat_id = " & CStr(rsAna!Pat_id) & Ausschluß
'         IF VorDat <> CDate(0) THEN
'             sql1 = sql1 + " AND zeitpunkt < " & DatFor_k(VorDat + 1)
'         Else
'             sql1 = "SELECT " & krzmd & "medikament, zeitpunkt FROM `rezepteintraege`  rz LEFT JOIN _lfaelle f ON rz.pat_id = f.pid WHERE zeitpunkt > SUBDATE(IF(f.bhfe1<now(),f.bhfe1,now()), INTERVAL 0.5 YEAR) AND f.pid = " & rsAna!Pat_id & Ausschluß
'         END IF
' '        SET rmed = Dtb.OpenRecordset(sql1, dbOpenDynaset)
'         rmed.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
''        SET rmA = TabÖff("medArten", "medikament")
'         IF Not rmed.BOF THEN
'          rmed.MoveFirst
'          Do While Not rmed.EOF
''           Kurzmed = UCase$(LEFT(rmed!Medikament, IIf(InStrB(Mid$(rmed!Medikament, 4), " ") > 0, InStr(Mid$(rmed!Medikament, 4), " ") - 1 + 4, Len(rmed!Medikament))))
''          rmA.Seek "=", Kurzmed
''           IF Not (LCase$(rmed!Medikament) LIKE "*senso*" OR LCase$(rmed!Medikament) LIKE "* ttr*" OR LCase$(rmed!Medikament) LIKE "*clix*" OR LCase$(rmed!Medikament) LIKE "*loe*" OR LCase$(rmed!Medikament) LIKE "*aviva*") THEN ' ACCU CHEK SENSOR COmF GLUC TTR 50 St / Accu Chek Spirit Serv Pack  1 P
'            SET rMA = Nothing
'            Dim MittelMed$
'            MittelMed = rmed!kurzmed
''            MittelMed = GetMed(rmed!Medikament, 0)
'            myFrag rMA, "SELECT puzu,InS,Anal FROM `medarten` WHERE medikament = '" & MittelMed & "'"
'            IF Not rMA.BOF THEN
'             IF NOT ISNULL(rMA!puzu) AND rMA!puzu THEN
'              obRezPumpe = True
'              Exit Do
'            ElseIf rMA!InS OR rMA!anal THEN
'              obRezIns = True
'             END IF
'            END IF
''           END IF
'           rmed.MoveNext
'          Loop
'          rmed.Close
'         END IF
'        END IF
'       END IF
'       IF InStrB(IIf(obanf, rsAna!Ther1, rsAna!TherAkt), "?") > 0 THEN therart_erm = offen
''       IF rsAna!j_Insulinpumpe <> 0 OR obRezPumpe THEN
'       IF rsAna!Insulinpumpe <> 0 OR obRezPumpe THEN
'         therart_erm = csii
'       Else
'         ON Error GoTo fehler
'         Dim ThF AS TherapieArt
'         ThF = TherAuskunft(rsAna!Pat_id, obanf, insz, VorDat, obIns, obAnal, obGlib, obmetf, obGlucI, obSHGlin, obGlit, obdpp4, obglp1, obsglt2, , , , , , , , Qmax, obRezIns, DT)
'         IF therart_erm < ThF THEN therart_erm = ThF
'       END IF
'  END SELECT
' Exit Function
'fehler:
' Dim AnwPfad$
'#If VBA6 THEN
' AnwPfad = CurrentDb.name
'#Else
' AnwPfad = App.path
'#END IF
' SELECT CASE MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(ISNULL(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in TherArt/" + AnwPfad)
'  Case vbAbort: Call MsgBox("Höre auf"): Progende
'  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
'  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
' END SELECT
'End FUNCTION ' Therart
#End If

Sub odbctest()
 Dim v1$(), v2$(), i
 Call regEnumVal("SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers", v1, v2)
 For i = 0 To UBound(v1)
  Debug.Print v1(i), v2(i)
 Next i
End Sub

Public Function bittest1()
 'SetDBCn Nothing
 DBCnS = vNS
 If LenB(DBCn) = 0 Or LenB(DBCnS) = 0 Then Call acon(quelleT)  ' DBCn.ConnectionString
 Dim rs As New ADODB.Recordset
 On Error Resume Next
 myFrag rs, "SELECT Bit FROM test.test"
 If rs.State <> 0 Then
  Debug.Print rs!Bit, DBCn
  rs.Move 1
  Debug.Print rs!Bit, DBCn
 End If ' rs.State <> 0 Then
End Function ' bittest1

Function bittest()
 Dim rAF&, Cn$(2), dtyp$(1), i%, j%
 Dim rs As New ADODB.Recordset
 Dim Vb As New ADODB.Connection
 On Error GoTo fehler
 Call Lese.ProgStart
 Cn(0) = "Provider=microsoft.Jet.OLEDb.4.0;Data Source=" & QmdB & ";Jet OLEDb:Engine Type=5" ' u:\anamnese\quelle.mdb
 Cn(1) = "PROVIDER=MSDASQL;driver={" & ODBCStr & "};server=linux1;uid=praxis;pwd=***REMOVED***;database=quelle;"
 Cn(2) = "PROVIDER=MSDASQL;driver={MySQL ODBC 3.51 Driver};server=linux1;uid=praxis;pwd=***REMOVED***;database=quelle;"
 dtyp(0) = "bit"
 dtyp(1) = "tinyint"
 For i = LBound(Cn) To UBound(Cn)
  For j = LBound(dtyp) To UBound(dtyp)
   Vb.Open Cn(i)
   On Error Resume Next
   myEFrag "DROP TABLE bittest", , Vb
   On Error GoTo fehler
   myEFrag "CREATE TABLE bittest (id " & IIf(InStrB(Cn(i), "MySQL") <> 0, "integer(10) auto_increment", "counter primary") & " key, ob " & dtyp(j) & ")", , Vb ' bit, bit(1), tinyint(1)
   myEFrag "INSERT INTO bittest(ob) VALUES(1)", , Vb
   myEFrag "INSERT INTO bittest(ob) VALUES(0)", , Vb
'   rs.Open "SELECT ob, --ob AS `minus minus ob`, id FROM bittest t", Vb, adOpenDynamic, adLockOptimistic
   myFrag rs, "SELECT ob, --ob AS `minus minus ob`, id FROM bittest t", adOpenDynamic, Vb
   Dim spezif$
   spezif = Vb.Provider
   On Error Resume Next
   spezif = spezif & ";V." & Vb.Properties("Driver ODBC Version")
   On Error GoTo fehler
   spezif = spezif & ":" & DefDB(Vb)
'   Debug.Print Left$(dtyp(j) & ":" & Space$(10), 10) & spezif & ":"
   Debug.Print Left$(dtyp(j) & ":" & Space$(10), 10) & Vb.Properties("Extended Properties")
   Do While Not rs.EOF
    Debug.Print rs.Fields(0).name & ": " & Left$(rs.Fields(0) & "," & Space$(10), 10) & rs.Fields(1).name & ": " & rs.Fields(1)
    rs.Move 1
   Loop
   rs.Close
   myEFrag "DROP TABLE bittest", , Vb
   Vb.Close
  Next j
 Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in bittest/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' bittest


Sub GibwerteAus(Optional obAnzeig As Boolean = True)
 do_GibwerteAus obAnzeig
' DiagnosenAusgeb ID, obanzeig
' DMPAusgeb0 ID, obanzeig
End Sub ' (Optional obAnzeig AS boolean = True)
Sub Formmerk()
' AktID = Forms(Anmnb)!Pat_id
 Call AnbogVar(True)
 AktID = Forms(Anmnbi)(ABPat_ID) 'Pat_IDfromForm
 altRecordSource = vNS
End Sub ' Formmerk()
Sub AnbogVar(Optional mitpat_id%)
 Dim i&, j&
 On Error Resume Next
 For i = 0 To Forms.COUNT - 1
  If Forms(i).name = Anmnb Then
   Anmnbi = i
   If mitpat_id Then
    For j = 0 To Forms(i).COUNT - 1
     Err.Clear
     If Forms(i)(j).DataField = "Pat_ID" Then
      If Err.Number = 0 Then
       ABPat_ID = j
       Exit Sub
      End If ' Err.Number = 0 THEN
     End If ' Forms(i)(j).DataField = "Pat_ID" THEN
    Next j
   End If ' mitpat_id
  End If ' Forms(i).Name = Anmnb THEN
 Next i
End Sub ' AnbogVar(Optional mitpat_id%)

Sub FormAufruf()
 On Error GoTo fehler
 syscmd acSysCmdSetStatus, "FormAufruf ..."
 If AktID <> 0 Then
  DoCmd.Close acForm, Anmnb, acSaveYes ' "Anamnesebogen nach Name"
'  DoCmd.OpenForm Anmnb '"Anamnesebogen nach Name" dtb.Containers(2).Documents(0).Name
  Call AnbogVar
  Forms(Anmnbi).Show
  DoCmd.Maximize
  Forms(Anmnbi).adoRS.Find " Pat_id >= " & CStr(AktID), 0, adSearchForward, adBookmarkFirst
'  ON Error Resume Next
'  Forms(anmnbi).Recordset.FindFirst "Pat_ID = " + CStr(AktID) ' wieder alten Datensatz wählen
  On Error GoTo fehler
  If Forms(Anmnbi).adoRS.EOF Then ' Recordset.Nomatch() THEN
   altRecordSource = Forms(Anmnbi).RecordSource
   Forms(Anmnbi).RecordSource = "Anamnesebogen alle Datensätze"
   On Error Resume Next
   Forms(Anmnbi).Recordset.FindFirst "ID = " + CStr(AktID) ' wieder alten Datensatz wählen
   On Error GoTo fehler
  End If
 End If
Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in FormAufruf/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Sub ' FormAufruf()

Sub FormRestoreSource()
 Dim i%
 Call AnbogVar
 If Forms(Anmnbi).name <> Anmnb Then
  DoCmd.Close acForm, Anmnb, acSaveYes
  DoCmd.OpenForm Anmnb
 End If
 If altRecordSource <> "" Then
  Forms(Anmnbi).adoRS.Close
'  Forms(Anmnbi).adoRS.Open altRecordSource, Lese.dbv.CnStr, adOpenDynamic, adLockOptimistic
  myFrag Forms(Anmnbi).adoRS, altRecordSource, adOpenDynamic, Lese.dbv.CnStr
 End If
End Sub ' FormRestoreSource()

Sub do_GibwerteAus(Optional obAnzeig As Boolean = True)
Dim ErgebDatei$
ErgebDatei$ = aVerz + "Anamnese.txt"
Dim te$
Call Formmerk
Call FormAufruf
te = machwertString$(CStr(AktID), obAnzeig)
Open ErgebDatei For Output Access Write Lock Write As #1
If te <> "" Then
 Print #1, te
End If
Close #1
Call FormRestoreSource
If obAnzeig Then zeigan ErgebDatei, vbNormalFocus
End Sub ' do_GibwerteAus(Optional obAnzeig AS boolean = True)

Function KommRep() ' Kommentare reparieren
 Dim Acc As New ADODB.Connection, MyS As New ADODB.Connection, ErrNo&
 Dim rs As New ADODB.Recordset, T1 As New ADODB.Recordset, T2 As New ADODB.Recordset, komm$
 Dim xc As New ADOX.Catalog, xr As New ADOX.Table
 Dim xcA As New ADOX.Catalog, xrA As New ADOX.Table
 Dim lauf As Variant
 Dim ErrNr&, ErrDes$
 On Error GoTo fehler
 Acc.Open "Provider=microsoft.Jet.OLEDb.4.0;Data Source=" & QmdB & ";Jet OLEDb:Engine Type=5"
 MyS.Open "PROVIDER=MSDASQL;driver={" & ODBCStr & "};server=linux1;uid=praxis;pwd=***REMOVED***;database=quelle;"
 myFrag rs, "SHOW TABLES", , MyS
 xc.ActiveConnection = MyS
 xcA.ActiveConnection = Acc
 Do While Not rs.EOF
  Set xr = xc.Tables(LCase$(rs.Fields(0)))
  If xr.Type = "TABLE" Then
'   Debug.Print rs.Fields(0)
   On Error Resume Next
'   Err.Clear
   ErrNo = -1
   If Not xrA Is Nothing Then
    For Each lauf In xcA.Tables
     If LCase$(lauf.name) = LCase$(rs.Fields(0)) Then
      Set xrA = xcA.Tables(lauf.name)
      ErrNo = 0
      Exit For
     End If
    Next lauf
   End If
'   SET T1 = Nothing
'   T1.Open "SELECT top 1 * FROM `" & rs.Fields(0) & "`", Acc
'   ErrNo = Err.Number
   On Error GoTo fehler
    If ErrNo = 0 Then
     'Set T2 = Nothing
     myFrag T2, "SHOW FULL COLUMNS FROM `" & rs.Fields(0) & "`", adOpenStatic, MyS
     Do While Not T2.EOF
      If T2!Comment = "" Then
       komm = vNS
       On Error Resume Next
       komm = xrA.Columns.Item(T2!Field.Value).Properties("Description")
'       komm = T1.Fields(T2!Field).Properties("Description")
       On Error GoTo fehler
       If komm <> T2!Comment Then
        Debug.Print "->", xr.name, T2!Field.Value, komm, T2!Comment
       End If
       If komm <> "" Then
        Dim AfN&
        On Error Resume Next
        Err.Clear
        Dim defa$
        Dim valu$
          If IsNull(T2!Default) Then
           defa = "NULL"
          Else
           defa = T2!Default
          End If
          valu = T2!Field.Value
        Call myEFrag("ALTER TABLE `" & xr.name & "` MODIFY COLUMN `" & T2!Field.Value & " " & T2!Type.Value & " DEFAULT " & defa & IIf(IsNull(T2!collation), vNS, " COLLATE " & T2!collation) & " COMMENT '" & komm & "', ENGINE = InnoDB;", AfN, MyS, True, ErrNr, ErrDes)
        If ErrNr <> 0 Then
         Err.Clear
         On Error GoTo fehler
         Debug.Print xr.name, T2!Field.Value, komm
         On Error Resume Next
         Err.Clear
         Call myEFrag("ALTER TABLE `" & xr.name & "` MODIFY COLUMN `" & T2!Field.Value & "` " & T2!Type.Value & " DEFAULT " & defa & IIf(IsNull(T2!collation), vNS, " COLLATE " & T2!collation) & " COMMENT '" & komm & "', ENGINE = InnoDB;", AfN, MyS, True, ErrNr, ErrDes)
         If ErrNr <> 0 Then
          On Error GoTo fehler
          Call myEFrag("ALTER TABLE `" & xr.name & "` MODIFY COLUMN `" & T2!Field.Value & "` " & T2!Type.Value & " DEFAULT " & defa & IIf(IsNull(T2!collation), vNS, " COLLATE " & T2!collation) & " COMMENT '" & komm & "', ENGINE = InnoDB;", AfN, MyS)
         End If
         On Error GoTo fehler
        End If
        On Error GoTo fehler
       End If
      End If
      T2.Move 1
     Loop
    End If
   End If
  rs.Move 1
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in kommrep/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' KommRep

Function AnReparieren() ' Ergänzt die in den mySQL-Tabellen vermutlich bei der Spaltenbreitenanpassung verlorenen Kommentare aus u:\anamnese\quelle.mdb
'ALTER TABLE `quelle1`.`anamnesebogen` mODIFY COLUMN `Grund für Vorstellung` VARCHAR(479) CHARACTER SET latin1 COLLATE latin1_german2_ci DEFAULT NULL COmmENT '^Grund für Vorstellung';
 Dim i&
 Dim DbA As New ADODB.Connection
 Dim dbm As New ADODB.Connection
 Dim rAn As New ADODB.Recordset
 Dim CStr1$
 On Error GoTo fehler
' Call doConstrFestleg(1, 1)
 Call aCStr(quelleT, accDtb)
 DbA.Open DBCnS ' DBCn.ConnectionString
 Dim k As ConDtb
 For k = qDtb To q2Dtb
'  Call doConstrFestleg(k, 1)
  CStr1 = aCStr(quelleT, k)
  Set rAn = Nothing
  myFrag rAn, "SELECT * FROM `anamnesebogen`", adOpenStatic, DbA, adLockReadOnly
  Set dbm = Nothing
  dbm.Open CStr1
  Dim myDbx As New ADOX.Catalog
  Dim myTable As ADOX.Table
  myDbx.ActiveConnection = DbA
  Set myTable = myDbx.Tables("anamnesebogen")
  For i = 0 To myTable.Columns.COUNT - 1
   Dim rsc As ADODB.Recordset
'   Set rsc = Nothing
'   Call rsc.Open("SHOW FULL COLUMNS FROM `" & Forms(0).MyDB & "`.`anamnesebogen` WHERE field = '" & rAn.Fields(i).name & "'", dbm, adOpenStatic, adLockReadOnly)
   myFrag rsc, "SHOW FULL COLUMNS FROM `" & Forms(0).MyDB & "`.`anamnesebogen` WHERE field = '" & rAn.Fields(i).name & "'", adOpenStatic, dbm, adLockReadOnly
   syscmd acSysCmdSetStatus, myTable.Columns(rAn.Fields(i).name).Properties("Description")
   Call myEFrag("ALTER TABLE `" & Forms(0).MyDB & "`.`anamnesebogen` modify COLUMN `" & rAn.Fields(i).name & "` " & rsc!Type & IIf(IsNull(rsc!collation), vNS, " CHARACTER SET utf8mb4 COLLATE " & rsc!collation) & " default " & IIf(IsNull(rsc!Default), " null", rsc!Default) & " comment '" & rsc!Comment & "'", , dbm) 'myTable.Columns(rAn.Fields(i).Name).Properties("Description") & "'")
  Next i
 Next k
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in machwertString/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AnReparieren()

Function obN%(fldv$)
 Select Case fldv
  Case "n", "0", "N", "-", "o.b", "o. b"
    obN = True
 End Select
End Function ' obN%

Function machwertString$(Pat_id$, Optional obAnzeig As Boolean = False) ' Anamnese für brief
Dim FNr&
Dim te$, Tü$, Add$, Tneu$, Descr$
Dim fld, fldv$
Dim obNz As Boolean ' neue zeile
Dim obkomma As Boolean ' Komma
Dim keinName As Boolean
Dim obVorwort As Boolean ' ob Vorwort exisitiert
Dim Vorwort$ ' Teil der überschrift für mehrere Tabellenfelder
Dim DPPos% ' Position des Doppelpunktes
Dim bmival#, stelle%
Dim j%, k%
Dim obnadp As Boolean
Dim condens$, rechts$, links$
#Const obmwSdao = 0
On Error GoTo fehler
#If obmwSdao Then
 Dim rAn As DAO.Recordset
 Set rAn = TabÖff("Anamnesebogen", "Pat_id") 'Forms(Anmnb).RecordsetClone
 rAn.Seek "=", CLng(Pat_id)
'ran.FindFirst ("Pat_ID = " + CStr(AktID)) 'CStr(ID))
 If rAn.Nomatch Then GoTo schluss
#Else
' IF lies Is Nothing THEN
 Call Lese.ProgStart
' END IF
 If Not lies.obMySQL Then
  Dim myDbx As New ADOX.Catalog
  Dim myTable As ADOX.Table
  myDbx.ActiveConnection = DBCn
  Set myTable = myDbx.Tables("anamnesebogen")
 End If
 Dim rAn As New ADODB.Recordset
' myFrag rAn, "SELECT -obbzausgew AS j_obbzausgew, -obosaufgek AS j_obosaufgek, -obpodaufgek AS j_obpodaufgek, -obmblausgeh AS j_obmblausgeh, -obschulaufgek AS j_obschulaufgek, -obdmpaufgekl AS j_obdmpaufgekl, -obmednetz AS j_obmednetz, -ob AS j_ob, -oban1eing AS j_oban1eing, -oban2eing AS j_oban2eing, -obanaeing AS j_obanaeing, -obcheck AS j_obcheck, -[Bypaß kardial] AS `j_bypass kardial` , -`bypaß peripher` AS `j_bypaß peripher`, -tkz AS j_tkz, -insulinpumpe AS j_insulinpumpe, -dialyse AS j_dialyse, a.* FROM `anamnesebogen` a WHERE pat_id = " & Pat_id
 myFrag rAn, "SELECT * FROM `anamnesebogen` WHERE pat_id = " & Pat_id, adOpenStatic
 If rAn.BOF Then GoTo schluss
#End If
te = vNS
Vorwort = vNS
For k = IIf(obAnzeig, 0, 7) To rAn.Fields.COUNT - 1
 Set fld = rAn.Fields(k)
 Descr = vNS
 On Error Resume Next
' Descr = Dtb.TableDefs(fld.SourceTable).Fields(fld.SourceField).Properties("Description")
 If lies.obMySQL Then
  Dim rsc As New ADODB.Recordset
  Set rsc = Nothing
  myFrag rsc, "SHOW FULL COLUMNS FROM `anamnesebogen` WHERE field = '" & rAn.Fields(k).name & "'", , , , , , , FNr
  Descr = rsc!Comment
 Else
  Descr = myTable.Columns(rAn.Fields(k).name).Properties("Description")
 End If
' FNr = Err.Number
 On Error GoTo fehler
 Err.Clear
 If FNr <> 0 Then ' Descr = Dtb.TableDefs(fld.SourceTable).Fields(fld.SourceField).Name
  Descr = rAn.Fields(k).name
 End If
 obNz = False
 If Left$(Descr, 1) = "^" Then
  obNz = True
  If te <> "" Then
'   te = ucase$(LEFT(te, 1)) + mid$(te, 2)
   If te <> "" Then
'    IF right$(te, 1) <> "." THEN te = te + "."
'    Print #1, te ' letzte Runde ausgeben
    If Right$(te, 1) <> vbLf Or te = "" Then te = te + IIf(Right$(te, 1) = ".", vNS, ".") & vbCrLf
   End If
'   te = vns
  End If
  Descr = LTrim$(Mid$(Descr, 2))
 End If
 
 obVorwort = False
 If Left$(Descr, 1) = "<" Then
  obVorwort = True
  Descr = LTrim$(Mid$(Descr, 2))
 End If
 
 obkomma = False
 If Left$(Descr, 1) = "," Then
  obkomma = True
  Descr = LTrim$(Mid$(Descr, 2))
 End If
 
 keinName = False
 If Left$(Descr, 1) = "-" Then
  keinName = True
  Descr = LTrim$(Mid$(Descr, 2))
 End If
 
 If Left$(Descr, 1) = ":" Then Descr = IIf(keinName, vNS, fld.name) + ":"
 DPPos = InStr(Descr, ":")
 If DPPos > 0 Then
  If Not obVorwort Then
   Vorwort = Left$(Descr, DPPos - 1) + ":"
   Descr = Vorwort & " " & Mid$(Descr, DPPos + 1)
  End If
 End If
 
 If obVorwort Then
  If Right$(te, 1) = vbLf Then
   Descr = Vorwort & " " & LTrim$(Descr)
  End If
 End If
 
 If obkomma Then Descr = ", " + Descr

 If Not IsNull(fld) Then
'  IF instrb(fldname, "chwanger") > 0 THEN
'    msgbox "Achtung: Feld Schwanger: " + fld
'  END IF
  If fld <> "False" And fld <> "" And Not ((fld.Type = adUnsignedTinyInt Or fld.Type = adTinyInt) And fld = "0") Then
   Add = vNS
   fldv = REPLACE$(fld.Value, "j ", "ja ")
   If fldv = "j" Then fldv = "ja"
   If ((fld.Type = adUnsignedTinyInt And fld.Value = 1) Or (fld.Type = adTinyInt And (fld.Value = 1 Or fld.Value = -1))) Then
    fldv = "ja"
   End If
   
   Dim fldName$
   fldName = LCase$(fld.name)
   Select Case fldName
    Case "aufschreiben", "bluthochdruck"
     fldv = REPLACE$(REPLACE$(fldv, "n ", "nein "), "n,", "nein,")
   End Select
   Select Case fldName
    Case "aufschreiben", "bluthochdruck", "bdselbst", "schwanger", "netzhaut gelasert", "sehminderung unbehebbar", "diabet nierenschaden", "albumin zuletzt", _
         "andere nierenerkrankung", "herzkrankheit", "angina pectoris", "herzinfarkt", "ptca oder stent", "herzschwäche", _
         "hirndurchblutungsstörung", "schlaganfall", "beindurchblutungsstörung", "schaufensterkrankheit", "geschwür", _
         "amputation", "ameisenlaufen", "druckstellen", "verformungen", "neue fußkomplikationen", "entleerungsstörungen magen", _
         "entleerungsstörungen harnblase", "schwindel aufstehen", "folgeerkrankungen haut", "bewegungseinschränkungen", _
         "sexualstörung", "alkohol", "tabak", "tabakex", "tabakbis", "tabakakt", "tabakmenge", "weitere medikation", "liphypertrophien abdomen", "liphypertrophien beine", _
         "liphypertrophien arme", "hyperkeratosen", "ulcera", "unterzucker pm", "keto pa", "bzgr300 pm", _
         "fußpflege", "podologie", "einlagen", "bzmessungen pw", "bzmessungen pw nde", "bzmessungen p w nachts", _
         "bzmessungen selbst", "aufschreiben", "beinödven" ' "schulung", "dmp",
     If fldv = "j" Or fldv = "J" Then
      Add = "ja"
     ElseIf fldv = "n" Or fldv = "0" Or fldv = "N" Or fldv = "-" Or fldv = "o.b" Or fldv = "o. b" Then
      Select Case fldName
       Case "unterzucker pm", "keto pa", "bzgr300 pm", "bzmessungen pw", "bzmessungen pw nde", "bzmessungen p w nachts"
        Add = "keine"
       Case Else
        Add = "nein"
      End Select
     ElseIf fldv = "u" Then
      Add = "unbekannt"
     Else
      If Left$(fldv, 2) = "n " Then
       Add = REPLACE$(fldv, "n ", "nein ", , 1)
      Else
       Add = CStr(fldv)
      End If
     End If
     fldv = Add
    Case "asr", "psr", "puls leiste", "puls kniekehle", "puls atp", "puls adp"
     Add = PulsParse(fldv)
     If InStrB(Add, "bi") <> 0 Or InStrB(Add, "mono") <> 0 Or InStrB(Add, "post") <> 0 Then
      Descr = REPLACE$(Descr, "Pulse", "Pulse/Dopplersignale")
     End If
     fldv = Add
   End Select
'   SELECT CASE fldname
'    Case "Oberflächensensibilität", "monofilamenttest", "Kalt-warm", "Vibration IK", "Vibration Großzehe"
'     j = InStr(fldV, "/")
'     IF j > 0 AND instrb(j + 1, fldV, "/") = 0 THEN
'      fldV = trim$(fldV) & ", " & trim$(fldV)
'     END IF
'   END SELECT
   Select Case fldName
    Case "liphypertrophien abdomen", "liphypertrophien beine", "liphypertrophien arme", "beinbefund", "hyperkeratosen", _
         "ulcera", "kraft zehenheber", "kraft zehenbeuger", "kraft knie", "asr", "psr", "oberflächensensibilität", _
         "puls leiste", "puls kniekehle", _
         "puls atp", "puls adp", "blutdruckwerte"
     If InStrB(fldv, ",") > 0 Then
      fldv = "rechts: " + LTrim$(Mid$(fldv, 1, InStr(fldv, ","))) + " links: " + LTrim$(Mid$(fldv, InStr(fldv, ",") + 1))
     End If
   End Select
   
   Select Case fldName
    Case "blutdruckwerte", "rr"
     fldv = REPLACE$(fldv, "arr.", "arrhythmisch")
     fldv = REPLACE$(fldv, "arrh.", "arrhythmisch")
     If InStr(fldv, "arrh") <> InStr(fldv, "arrhythmisch") Then
      fldv = REPLACE$(fldv, "arrh", "arrhythmisch")
     ElseIf InStr(fldv, "arr") <> InStr(fldv, "arrhythmisch") Then
      fldv = REPLACE$(fldv, "arr", "arrhythmisch")
     End If
    Case "herz", "carotiden"
     fldv = REPLACE$(fldv, "SG", "Strömungsgeräusch")
   End Select
   obnadp = False
   j = InStr(1, fldv, "nadp", vbTextCompare)
   If j > 0 Then
    obnadp = True
    fldv = Left$(fldv, j - 1) + LTrim$(Right$(fldv, Len(fldv) - j - 3))
   End If
   
   Select Case fldName
    Case "gebdat", "tkz", "j_tkz", "pat_id", "hanr", "hanr2", "letzte änderung", "diagnosen", "vorgestellt", "schulung", _
         "dmp", "dmschl", "rrschulz", "vorname", "nvorsatz", "titel", "anrede", "rrturbomed", "versicherung", _
         "dmschulz", "dmschl", "rrschulz", "dmphier", "aktzeit", "ther1", "therakt", "prim", "oban1eing", "j_oban1eing", _
         "j_oban2eing", "oban2eing", "j_obanaeing", "obanaeing", "j_obcheck", "obcheck", "j_obbzausgew", "obbzausgew", "j_obosaufgek", "obosaufgek", "j_obpodaufgek", "obpodaufgek", "j_obmblausgeh", "obmblausgeh", _
         "j_obschulaufgek", "obschulaufgek", "j_obdmpaufgekl", "obdmpaufgekl", "j_obmednetz", "obmednetz", "j_ob", "ob", "versicherungsart", "hausarzt", "qs", "qt"
    Case "diabetestyp"
     Add = " " + dtyp(rAn.Fields(fldName))
    Case "nachname"
     Add = CStr(fldv) & ", " & rAn!Vorname
     If Not IsNull(rAn!Titel) Then
      If Len(rAn!Titel) > 0 Then
       Add = rAn!Titel & " " & Add
      End If
     End If
     If Not IsNull(rAn!anrede) Then
      If Len(rAn!anrede) > 0 Then
       Add = rAn!anrede & " " & Add
      End If
     End If
    Case "diabetes seit"
     If fldv = "bu" Or fldv = "b" Then
      Descr = ", "
      Add = "bisher unbekannt"
     Else
      Add = CStr(fldv)
     End If
    Case "weitere anamnese"
     Add = CStr(fldv)
    Case "tendenz"
     Select Case fldv
      Case "a"
       Add = "abnehmend"
      Case "g"
       Add = "gleichbleibend"
      Case "z"
       Add = "zunehmend"
     End Select
    Case "tabletten seit", "insulin seit"
     If fldv = "-" Then Add = vNS Else Add = CStr(fldv)
    Case "j_insulinpumpe", "insulinpumpe"
     If fldv <> "" Then Add = "ja"
    Case "größe", "gewicht"
     If fldv = "0" Then Add = vNS Else Add = CStr(fldv)
     If fldName = "gewicht" Then
      If IsNumeric(Add) Then If CDbl(Add) < 3 Then Add = CStr(CDbl(Add) * 100)
      Add = Add + " kg"
      On Error Resume Next
      bmival = 0
      bmival = rAn!Gewicht / rAn!Größe / rAn!Größe '* 10000
      Do
       If bmival = 0 Then Exit Do
       If bmival > 8 Then Exit Do
       bmival = bmival * 10
      Loop
      If bmival <> 0 Then
       Add = Add + ", BMI = " + Format$(bmival, "###.#") + " kg/m²"
      End If
      On Error GoTo fehler
     Else ' fldname = "Größe"
      Add = Add + " cm"
     End If
    Case "essenszeit früh", "essenszeit mittags", "essenszeit abends", "essenszeit spät"
     If fldv = "-" Then
      Add = "-"
     Else
      Add = CStr(fldv) + " Uhr"
     End If
    Case "spritz-eß-abstand früh", "spritz-eß-abstand mittags", "spritz-eß-abstand abends"
     Add = CStr(fldv) + " min"
    Case "spritzstelle früh", "spritzstelle mittags", "spritzstelle abends", "spritzstelle nachts"
     Add = Spritzstelle(fldv)
'    Case "Puls Leiste"
'     Add = "Leiste " + CStr(fldV)
'    Case "Liphypertrophien Abdomen"
'     Add = "Abdomen " + CStr(fldV)
    Case "uz rechtzeitig"
     Select Case fldv
      Case "i"
       Add = "immer"
      Case "m"
       Add = "meist"
      Case "n"
       Add = "nie"
      Case Else
       Add = REPLACE$(REPLACE$(REPLACE$(fldv, "j ", "ja "), "m ", "meist "), "n ", "nie ")
       If Left$(Add, 2) = "i " Then Add = REPLACE$(Add, "i ", "immer ", , 1)
     End Select
    Case "letztes hba1c", "vorherige werte"
     Add = fldv + " %"
     If InStr("0123456789", Right$(Trim$(fldv), 1)) > 0 And LenB(Trim$(fldv)) <> 0 Then
      Add = CStr(fldv) + " %"
     Else
      Add = CStr(fldv)
     End If
    Case "gemessen am"
     Add = datUmwandel(fldv)
     On Error Resume Next
     If Day(CDate(fldv)) = 1 Then
      Add = "ca. " + Add
     End If
     On Error GoTo fehler
    Case "fremde hilfe pa", "bewußtlos pa"
     If fldv = "0" Or fldv = "n" Or fldv = "-" Then
       Add = "kein mal"
     ElseIf Left$(fldv, 1) = "(" Then
      Add = fldv
     Else
      Add = fldv + " mal"
     End If
    Case "blutdruckwerte", "rr"
'     IF instrb("0123456789", right$(trim$(fldV), 1)) > 0 THEN
'      Add = CStr(fldV) + " mm Hg"
'     Else
       Add = CStr(fldv)
'     END IF
     stelle = InStr(Add, "/")
'     IF stelle = 0 THEN
'      IF instrb(fldV, "mm Hg") = 0 THEN
'       Add = CStr(fldV) + " mm Hg"
'      END IF
'     END IF
     While stelle <> 0
      stelle = stelle + 1
      While InStrB("0123456789", Mid$(Add, stelle, 1)) > 0 And stelle <= Len(Add)
       stelle = stelle + 1
      Wend
      If Mid$(Add, stelle, 2) <> "mm" And Mid$(Add, stelle + 1, 2) <> "mm" Then
       Add = Left$(Add, stelle - 1) + " mm Hg" + Right$(Add, Len(Add) - stelle + 1)
      End If
      stelle = InStr(stelle + 1, Add, "/")
     Wend
     If Trim$(LCase$(Add)) = "n" Then Add = "normal"
    Case "tabak", "tabakex", "tabakbis", "tabakakt"
     Add = CStr(fldv)
    Case "tabakmenge"
     If InStrB("0123456789", Right$(Trim$(fldv), 1)) > 0 And Not IsNumeric(Right$(Trim$(fldv), 4)) And Trim$(fldv) <> "" Then
      Add = CStr(fldv) + " zig./d"
     Else
      Add = CStr(fldv)
     End If
    Case "bzwerte v d essen", "bzwerte n d essen"
     If InStrB("0123456789", Right$(Trim$(fldv), 1)) > 0 And LenB(Trim$(fldv)) <> 0 Then
      Add = CStr(fldv) + " mg/dl"
     Else
      Add = CStr(fldv)
     End If
    Case "jahr letzte diabetesschulung"
     If fldv = "-" Then
      Add = "keine"
     Else
      Add = CStr(fldv)
     End If
    Case "ort schulung"
     If fldv = "-" And rAn![Jahr letzte Diabetesschulung] = "-" Then
      fldv = vNS
     Else
      Select Case UCase$(fldv)
       Case "KD"
        Add = "Klinikum Dachau"
       Case "KMB"
        Add = "Klinikum bogenhausen"
       Case "KMS"
        Add = "Klinikum Schwabing"
       Case Else
        Add = CStr(fldv)
      End Select
      If InStrB(Add, "linikum") > 0 Or InStrB(Add, "rankenh") > 0 Then
       Descr = "im"
      ElseIf InStrB(Add, "Dr.") > 0 Then
       Descr = "bei"
      Else
       Descr = "in"
      End If
     End If
    Case "augensp befund", "herz", "lunge", "bauch", "mundhöhle", "carotiden"
     If LCase$(fldv) = "n" Or LCase$(fldv) = "ob" Or LCase$(fldv) = "ok" Or LCase$(Left$(fldv, 3)) = "o.b" Then
      Select Case fldName
       Case "carotiden"
        Add = "kein Strömungsgeräusch"
       Case Else
        Add = "unauffällig"
      End Select
     ElseIf fldv = "neg" Then
      Add = "negativ"
     Else: Add = CStr(fldv)
     End If
    Case "erhöht?"
     If Trim$(fldv) <> "" Then
      If InStrB("n-nein", LCase$(fldv)) = 0 Then
       Add = "normal"
      ElseIf InStrB("j+ja", LCase$(fldv)) = 0 Then
       Add = "erhöht"
      End If
     End If
    Case "j_dialyse", "j_bypass kardial", "j_bypaß peripher", "dialyse", "bypass kardial", "bypaß peripher"
     If obN(fldv) Then Add = vNS Else Add = "ja"
'    Case "dialyse", "bypass kardial", "bypaß peripher"
'     IF lcase(fldv) = "falsch" THEN
'      Add = vns
'     ElseIf lcase(fldv) = "wahr" THEN
'      Add = "ja"
'     END IF
    Case "herzinfarkt wann", "verformungen beschreibung"
     If obN(fldv) Then
      Add = vNS
     Else
      Add = CStr(fldv)
     End If
    Case "kraft zehenheber", "kraft zehenbeuger", "kraft knie"
     If fldv = "n" Then
      Add = "seitengleich normal"
     ElseIf fldv = "sg" Or fldv = "sehr gut" Or fldv = "sehr kräftig" Then
      Add = "seitengleich sehr kräftig"
     Else
      Add = CStr(fldv)
     End If
    Case "oberflächensensibilität"
     If fldv = "n" Then
      Add = "seitengleich ungestört"
     ElseIf fldv = "?" Then
      Add = "?/?"
     Else
      Add = CStr(fldv)
     End If
    Case "monofilamenttest", "kalt-warm", "vibration ik", "vibration großzehe"
      Call bruchteile(fldv, rechts, links)
      If IsNumeric(Right$(rechts, 1)) Then
       If InStrB(rechts, "/") = 0 Then
        Select Case fld.name
         Case "monofilanmenttest", "kalt-warm"
          rechts = rechts + "/5"
         Case Else
          rechts = rechts + "/8"
        End Select
       End If
      End If
      If IsNumeric(Right$(links, 1)) Then
       If InStrB(links, "/") = 0 Then
        Select Case fldName
         Case "monofilanmenttest", "kalt-warm"
          links = links + "/5"
         Case Else
          links = links + "/8"
        End Select
       End If
      End If
      If REPLACE$(REPLACE$(REPLACE$(rechts, "rechts", ""), "re ", ""), "re.", "") = REPLACE$(REPLACE$(REPLACE$(links, "links", ""), "li ", ""), "li.", "") Then
        Add = IIf(InStr(rechts, "bds"), vNS, "bds. ") + REPLACE$(REPLACE$(REPLACE$(rechts, "rechts", ""), "re ", " "), "re.", "")
      Else
       If InStrB(rechts, "rechts") = 0 And InStrB(rechts, "re ") = 0 And InStrB(rechts, "re.") = 0 Then
        rechts = "re: " + rechts
       End If
       If InStrB(links, "links") = 0 And InStrB(links, "li ") = 0 And InStrB(links, "li.") = 0 Then
        links = "li: " + links
       End If
       Add = rechts & ", " & links
      End If
    Case "ws", "nl", "nnh"
     If fldv = "n" Then
      Add = "kein Klopfschmerz"
     Else: Add = REPLACE$(fldv, "KS", "Klopfschmerz")
     End If
    Case "zähne"
     Add = REPLACE$(REPLACE$(fldv, "TP", "Teilprothese"), "VP", "Vollprothese")
    Case "sd"
     If fldv = "n" Then
      Add = "nicht vergrößert tastbar"
     Else: Add = CStr(fldv)
     End If
    Case "neuro sonst"
     Add = REPLACE$(REPLACE$(fldv, "KR", "Konvergenzreaktion"), "LR", "Lichtreaktion")
    Case "bmi"
     If bmival = 0 Then
      Add = fldv & " kg/m²"
     Else
      Add = ""
     End If
    Case Else
     Add = CStr(fldv)
   End Select
   If obnadp Then Add = "nach Angabe d.Pat. " + Add
   If obNz And Add = vNS And fldName <> "j_insulinpumpe" And fldName <> "insulinpumpe" Then Add = "-"
' 2. zeile Ende gestrichen: 'IIf(Descr = ":", fldname, "") +
   If Add <> vNS Then _
    te = te + IIf(Right$(te, 1) = vbLf Or LenB(te) = 0, vNS, IIf(Left$(Descr, 1) = ",", vNS, " ")) + _
         Trim$(IIf(Right$(te, 1) = vbLf Or LenB(te) = 0, IIf(Left$(Descr, 2) = ", ", Mid$(Descr, 3), Descr), Descr)) + IIf(LenB(te) = 0, vNS, " ") + Add
  End If ' fld <> "False" AND fld <> "" THEN
 End If ' NOT ISNULL(fld)
Next
te = UCase$(Left$(te, 1)) + Mid$(te, 2)
If te <> "" Then
 If Right$(te, 3) <> "." & vbCrLf Then te = Left$(te, Len(te) - 2) + "." & vbCrLf
End If
te = te + myEFrag("SELECT COALESCE(group_concat(concat(case when art LIKE 'Alk%' THEN 'Alkhol' when art LIKE 'fa%' AND art<>'fams' THEN 'Familienanamnese' when art LIKE 'rauch%' OR art LIKE 'nik%' THEN 'Tabak' end,' (',DATE_FORMAT(zeitpunkt,'%d.%m.%y'),'):','\t', Inhalt,'\n') SEPARATOR ''),'') FROM eintraege WHERE pat_id =" & Pat_id & "").Fields(0)
'  AND (art LIKE 'fa%' OR art LIKE 'rauch%' OR art LIKE 'nik%' OR art LIKE 'alk%') AND art<>'fams' ' braucht's offenbar nicht
machwertString = te
schluss:
rAn.Close
Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vNS, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in machwertString/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): ProgEnde
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'machwertString

Function bruchteile(fldv$, rE$, li$)
      Dim tp%  ' Trennposition
      Dim tp1% ' Trennposition 1
      tp = InStr(fldv, "|")
      If InStrB(fldv, ",") > 0 And InStrB(fldv, "bds") = 0 Then
       rE = Left$(fldv, InStr(fldv, ",") - 1)
       li = Right$(fldv, Len(fldv) - InStr(fldv, ","))
      ElseIf tp > 0 Then
       rE = Left$(fldv, tp - 1)
       li = Mid$(fldv, tp + 1)
       tp = InStr(rE, "/")
       If tp > 0 Then rE = Left$(rE, tp - 1)
       tp = InStr(li, "/")
       If tp > 0 Then li = Left$(li, tp - 1)
      Else
       rE = fldv
       li = fldv
      End If
      rE = LTrim$(Trim$(rE))
      li = LTrim$(Trim$(li))
End Function ' bruchteile(fldv$, re$, li$)
Function Spritzstelle(Abkü)
  Select Case LCase$(CStr(Abkü))
   Case "-"
    Spritzstelle = "-"
   Case "b"
    Spritzstelle = "Bauch"
   Case "o", "os"
    Spritzstelle = "Oberschenkel"
   Case "a"
    Spritzstelle = "Arm"
   Case "meist a"
    Spritzstelle = "meist Arm"
   Case "meist b"
    Spritzstelle = "meist bauch"
   Case "meist os"
    Spritzstelle = "meist Oberschenkel"
   Case "b/os", "os/b"
    Spritzstelle = "wechselnd bauch/Oberschenkel"
   Case "b/o/a"
    Spritzstelle = "wechselnd bauch/Oberschenkel/Arm"
   Case "b os"
    Spritzstelle = "bauch bzw. Oberschenkel"
   Case "a/b", "b/a"
    Spritzstelle = "wechselnd Arm/bauch"
  End Select
End Function ' Spritzstelle(Abkü)


