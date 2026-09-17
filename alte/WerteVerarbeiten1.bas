Attribute VB_Name = "WerteVerarbeiten"
Option Explicit
Public Const artspez$ = "(art in (""notiz"",""ni"",""telef"",""gs"",""rz"",""ga"",""ag"",""hz"",""ts"",""cr"",""ep"")"
Dim DMPStr$
Dim rDT As DAO.Recordset, sstr$
Dim Pat_ID&
'Enum DFSNiveau
'   stNichts%
'   St0%
'   St1%
'   St2%
'   AmpZeh%
'   St3%
'   AmpUS%
'   St4%
'   AmpOS%
'  End Enum

Sub WerteAnzeig()
Dim Abfr
Dim i&, runde%
Dim bisher()
Dim BCt%, aktBCt%
Const Dat0$ = aVerz & "AlleWerte.txt"
Const Dat1$ = aVerz & "VerschiedeneWerte.txt"
Dim dat(2)
dat(0) = Dat0
dat(1) = Dat1
For runde = 0 To 1
 Set Abfr = Dtb.OpenRecordset(Anmnb, dbOpenDynaset)
 If Not Abfr.BOF Then
  Open dat(runde) For Output Access Write Lock Write As #1
  For i = 1 To Abfr.Fields.Count - 1
   Print #1, Abfr.Fields(i).Name
   Abfr.MoveFirst
   BCt = 0
   ReDim bisher(0)
   While Not Abfr.EOF
    If Not IsNull(Abfr.Fields(i).Value) Then
     If Not Abfr.Fields(i).Value = "" Then
      If runde = 1 Then
       For aktBCt = 0 To BCt - 1
        If bisher(aktBCt) = Abfr.Fields(i).Value Then GoTo gibts
       Next
       bisher(BCt) = Abfr.Fields(i).Value
       BCt = BCt + 1
       ReDim Preserve bisher(BCt)
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
 Call Shell("notepad " + dat(runde), vbNormalFocus)
Next runde
End Sub ' WerteAnzeig()
Sub WAusgeb()
WerteAusGeb True
End Sub ' WAusgeb()
Sub WerteAusGeb(Optional obstumm As Boolean = False)
'Dim ID$
'Dim tbl
'' DoCmd.Save acForm, Anmnb geht leider nicht
'On Error Resume Next
'ID = forms!Anamnesebogen!Pat_ID
'On Error GoTo 0
'If ID = "" Then
' For Each tbl In dtb.TableDefs
'  DoCmd.Close acTable, tbl.Name, acSaveYes
' Next
' DoCmd.OpenForm dtb.Containers(2).Documents(0).Name 'Anamnesebogen
' DoCmd.Maximize
' ID = forms!Anamnesebogen!Pat_ID
'End If
Dim rs As DAO.Recordset
Call FormMerk
Call FormAufruf
Call GibWerteAus(Not obstumm)
Call DiagnosenAusgeb(AktID, Not obstumm)
Set rs = TabÖff("Anamnesebogen", "Pat_id")
rs.Seek "=", AktID
If rs.NoMatch Or (Not rs.NoMatch And (rs!Diabetestyp = "2" Or rs!Diabetestyp = "s")) Then
 Call DMPAusgeb(AktID, Not obstumm)
End If
Call BehDatAusgeb(AktID, Not obstumm)
Call LaborAusgeb(AktID)
Call FormAufruf
End Sub ' WerteAusGeb
Sub DiagnosenAusgeb(Optional ID&, Optional obAnzeig As Boolean = True)
Const ErgebDatei$ = aVerz + "Diagnosen.txt"
'Dim EStr$
Dim tbl
' DoCmd.Save acForm, Anmnb geht leider nicht
If ID = 0 Then
 On Error Resume Next
 ID = Forms!Anamnesebogen!Pat_ID
 On Error GoTo 0
 If ID = 0 Then
  For Each tbl In Dtb.TableDefs
   DoCmd.Close acTable, tbl.Name, acSaveYes
  Next
  DoCmd.OpenForm Dtb.Containers(2).Documents(0).Name
  DoCmd.Maximize
  ID = Forms!Anamnesebogen!Pat_ID
 End If
End If
Open ErgebDatei For Output Access Write Lock Write As #1
'EStr = Replace(Forms!Anamnesebogen!Diagnosen, " [", Chr(9) + "[")
'EStr = Left(EStr, Len(EStr) - 2)
Print #1, AusgDiag(ID)
Close #1
If obAnzeig Then Call Shell("notepad " + ErgebDatei, vbNormalFocus)
End Sub ' Diagnosenausgeb
Function AusgDiag$(Pat_ID&, Optional ohneNotwend%)
 Dim rAna As DAO.Recordset
 Dim spl$(), j%
 Set rAna = TabÖff("Anamnesebogen", "Pat_ID")
 rAna.Seek "=", Pat_ID
 If Not rAna.NoMatch Then
  spl = Split(rAna!Diagnosen, Chr(11))
  For j = 0 To UBound(spl)
   If ohneNotwend Then If InStr(spl(j), "Notw") > 0 Then spl(j) = ""
   spl(j) = Replace(spl(j), " [", Chr(9) + "[")
   If Not obkNeph Or InStr(spl(j), "N08.3") = 0 Then ' Nephropathie ggf. streichen
    If spl(j) <> "" Then AusgDiag = AusgDiag + spl(j) + Chr(11)
   End If
  Next j
'  AusgDiag = Replace(rAna!Diagnosen, " [", Chr(9) + "[")
  If Len(AusgDiag) >= 1 Then AusgDiag = Left(AusgDiag, Len(AusgDiag) - 1)
 End If
 rAna.Close
End Function ' AusDiag
Sub BehDatAusgeb(Optional ID&, Optional obAnzeig As Boolean = True)
Dim lBehD As Date, tbl As TableDef
Dim rNo As DAO.Recordset
Const ErgebDatei$ = aVerz + "BehDat.txt"
If ID = 0 Then
On Error Resume Next
ID = Forms!Anamnesebogen!Pat_ID
On Error GoTo 0
If ID = 0 Then
 For Each tbl In Dtb.TableDefs
  DoCmd.Close acTable, tbl.Name, acSaveYes
 Next
 DoCmd.OpenForm Dtb.Containers(2).Documents(0).Name
 DoCmd.Maximize
 ID = Forms!Anamnesebogen!Pat_ID
End If
End If
Pat_ID = ID
Open ErgebDatei For Output Access Write Lock Write As #2
lBehD = leBe(Pat_ID)
Print #2, CStr(Forms!Anamnesebogen!Vorgestellt) + IIf(lBehDat <> CDate(0), ", zuletzt am " + Format(lBehD, "dd/mm/yy") + ",", "")
Close #2
End Sub ' BehDatAusgeb
Function erBe(ByVal Pat_ID&) As Date
Static rFa As DAO.Recordset
'Set rFa = TabÖff("faelle", "ErstF")
Dim rFaName$
On Error Resume Next
rFaName = rFa.Name
On Error GoTo 0
If rFaName = "" Then Set rFa = TabÖff("faelle", "Auswahl")
rFa.Seek "=", Pat_ID
' erBe = IIf(rFa!SchGr = "90", rFa!BhFB, IIf(rFa!ausgst = CDate(0), rFa!lVorl, rFa!ausgst))
Do
 If rFa.EOF Or rFa.NoMatch Then Exit Do
 If rFa!Pat_ID <> Pat_ID Then Exit Do
 If IsNull(rFa!Fanf) Then Call doFAnfFuell(rFa)
 If rFa!Fanf < erBe Or erBe = 0 Then erBe = rFa!Fanf
 rFa.Move 1
Loop
End Function ' erBe(ByVal Pat_id&) As Date
Function leBe(ByVal Pat_ID&) As Date
Dim lBehD As Date, rLB As DAO.Recordset
On Error Resume Next
leBe = erBe(Pat_ID)
lBehD = Dtb.OpenRecordset("select max(zeitpunkt) as zp from [" + QMdbAkt + "].MedPlan where pat_id = " + CStr(Pat_ID), dbOpenDynaset)!Zp
If lBehD > leBe Then leBe = lBehD
lBehD = Dtb.OpenRecordset("select max(zeitpunkt) as zp from [" + QMdbAkt + "].vTextB(180) where " + artspez + " or left(art,2)=""rr"" or left(art,2)=""bz"") and pat_id = " + CStr(Pat_ID), dbOpenDynaset)!Zp
If lBehD > leBe Then leBe = lBehD
lBehDat = lBehD
End Function ' letzt
Function BehDauerStr$(Pat_ID)
 Dim d1 As Date, d2 As Date
 d1 = DateValue(erBe(Pat_ID))
 d2 = DateValue(leBe(Pat_ID))
 If d1 = d2 Then
  BehDauerStr = "am " + Format(d1, "d.M.YY")
 Else
  BehDauerStr = "vom " + Format(d1, "d.M.YY") + " bis zum " + Format(d2, "d.M.YY") + " mehrmals"
 End If
End Function ' BehDauerStr$()
Sub LaborAusgeb(Optional ID&)
 Const ErgebDatei$ = aVerz + "Labor.txt"
 If ID = 0 Then
  On Error Resume Next
  ID = Forms!Anamnesebogen!Pat_ID
  On Error GoTo 0
 End If
 Close #15
 Open ErgebDatei For Output Access Write Lock Write As #15
 Print #15, LaborString(ID)
 Close #15
End Sub
 
Function LaborString$(Pat_ID&)
 Dim rLW As DAO.Recordset, rDat As DAO.Recordset, ls$, rLU As DAO.Recordset
' Zeilenzahl bestimmen
 Dim zZ&, rZ$, gschl$, vgl$, altGruppe%, NB$ ' Normbereich
 Dim u!, o! ' oberer und unterer Grenzwert numerisch
 Dim uNG$, oNG$ ' obere und untere Normgrenze in Zeichen
 Dim pKz$ ' Pathologisch-Kennzeichen
 Dim abküZ%, gruppeZ%, zpZ%
 'dtbInit
 sql = "select Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" as NB from labor where pat_id = " + CStr(Pat_ID) + " UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar,Normbereich as NB " + _
  "FROM LaborXUS LEFT JOIN LaborXWert ON [LaborXUS].[RefNr]=[LaborXWert].[RefNr] " + _
  "WHERE pat_id = " + CStr(Pat_ID) + " and not exists (select * from laborNeu where pat_id = " + CStr(Pat_ID) + " and abkü = LaborXWert!Abkü and wert = LaborXWert.wert and zeitpunkt > LaborXUS.Eingang-3 and zeitpunkt < LaborXUS.Eingang+6)"
 Set rLU = Dtb.OpenRecordset(sql)
 gschl = ""
 altGruppe = 0
 Set rLW = TabÖff("Namen", "Pat_id")
 rLW.Seek "=", Pat_ID
 If Not rLW.NoMatch Then
  gschl = rLW!Geschlecht
 End If
 rLW.Close
 sql1 = "select count(*) from (select distinct abkü from (" + sql + "))"
 Set rLW = Dtb.OpenRecordset(sql1)
' Set rLW = Dtb.OpenRecordset("select count(*) from (SELECT distinct abkü  from [" + QMdbAkt + "].laborunion where pat_id = " + CStr(Pat_id) + ")")
 zZ = rLW.Fields(0)
 rLW.Close
 sql1 = "select count(*) from (select distinct gruppe from (SELECT * from (" + sql + ") as sql left join [" + QMdbAkt + "].laborparameter on sql.abkü = laborparameter.abkü where pat_id = " + CStr(Pat_ID) + "))"
 Set rLW = Dtb.OpenRecordset(sql1)
' Set rLW = Dtb.OpenRecordset("select count(*) from (select distinct gruppe from (SELECT * from [" + QMdbAkt + "].LaborUnion left join [" + QMdbAkt + "].laborparameter on LaborUnion.abkü = laborparameter.abkü where pat_id = " + CStr(Pat_id) + "))")
 zZ = zZ + rLW.Fields(0)
 rLW.Close
 sql1 = "select count(*) from (SELECT distinct(cdate(int(zeitpunkt)))  from [" + QMdbAkt + "].labor left join [" + QMdbAkt + "].laborparameter on Labor.abkü = laborparameter.abkü where pat_id = " + CStr(Pat_ID) + ")"
 Set rDat = Dtb.OpenRecordset(sql1)
' Set rdat = Dtb.OpenRecordset("select count(*) from (SELECT distinct(cdate(int(zeitpunkt)))  from [" + QMdbAkt + "].labor left join [" + QMdbAkt + "].laborparameter on Labor.abkü = laborparameter.abkü where pat_id = " + CStr(Pat_id) + ")")
 rZ = rDat.Fields(0)
 rDat.Close
 'statt datserial
 sql1 = "select * from (SELECT distinct(cdate(int(zeitpunkt))) as Datum from (" + sql + ") as sql left join [" + QMdbAkt + "].laborparameter on sql.abkü = laborparameter.abkü where pat_id = " + CStr(Pat_ID) + ") order by datum"
 Set rDat = Dtb.OpenRecordset(sql1)
' Set rdat = Dtb.OpenRecordset("select * from (SELECT distinct(cdate(int(zeitpunkt))) as Datum from [" + QMdbAkt + "].LaborUnion left join [" + QMdbAkt + "].laborparameter on LaborUnion.abkü = laborparameter.abkü where pat_id = " + CStr(Pat_id) + ") order by datum")
 If rDat.BOF Then Exit Function
 
 ls = CStr(rZ) + Chr(13) + Chr(10)
 rDat.MoveFirst
 Do While Not rDat.EOF
  ls = ls + Format(rDat!Datum, "dd.mm.yyyy") + Chr(13) + Chr(10)
  rDat.Move 1
 Loop
 
 ls = ls + CStr(zZ) + Chr(13) + Chr(10)
 sql1 = "select * from (SELECT *,datevalue(zeitpunkt) as Datum  from (" + sql + ") as sql left join [" + QMdbAkt + "].laborparameter on sql.abkü = laborparameter.abkü where pat_id = " + CStr(Pat_ID) + ") order by gruppe,reihe,datum"
 Set rLW = Dtb.OpenRecordset(sql1)
' Set rLW = Dtb.OpenRecordset("select * from (SELECT *,datevalue(zeitpunkt) as Datum  from [" + QMdbAkt + "].LaborUnion left join [" + QMdbAkt + "].laborparameter on LaborUnion.abkü = laborparameter.abkü where pat_id = " + CStr(Pat_id) + ") order by gruppe,reihe,datum")
 rLW.MoveFirst
 Do While Not rLW.EOF
   If Not IsNull(rLW!gruppe) Then
    If Val(rLW!gruppe) <> altGruppe Then
     ls = ls + Chr(13) + Chr(10)
    End If
   End If
   ls = ls + rLW.Fields("sql.LangText") + Chr(13) + Chr(10)
   ls = ls + nz(rLW.Fields("sql.einheit"),"") + Chr(13) + Chr(10)
   ls = ls + Format(rLW!Datum, "dd.mm.yyyy") + Chr(13) + Chr(10)
   ls = ls + nz(rLW!Wert,"") + Chr(13) + Chr(10)
   u = -99
   o = -99
   pKz = "n" ' normal
   uNG = ""
   oNG = ""
   If gschl = "w" Then
    If (Not IsNull(rLW!unw) And rLW!unw <> "") Then
     uNG = Replace(rLW!unw, "1:", "")
    End If
    If (Not IsNull(rLW!onw) And rLW!onw <> "") Then
     oNG = Replace(rLW!onw, "1:", "")
    End If
   End If
   If uNG = "" And oNG = "" Then
    If Not IsNull(rLW!unm) Then uNG = Replace(rLW!unm, "1:", "")
    If Not IsNull(rLW!onm) Then oNG = Replace(rLW!onm, "1:", "")
   End If
   If uNG <> "" Then u = CDbl(uNG)
   If oNG <> "" Then o = CDbl(oNG)
'   If IsNumeric(rLw!Wert) Or rLw!Wert Like "1:*" Then
'    If gschl = "w" Then
'     If Not rLw!unw = "" Then u = CDbl(Replace(rLw!unw, "1:", ""))
'     If Not rLw!onw = "" Then o = CDbl(Replace(rLw!onw, "1:", ""))
'    End If
'    If gschl = "m" Or u = -99 Then
'     If Not rLw!unm = "" Then u = CDbl(Replace(rLw!unm, "1:", ""))
'    End If
'    If gschl = "m" Or o = -99 Then
'     If Not rLw!onm = "" Then o = CDbl(Replace(rLw!onm, "1:", ""))
'    End If
   If IsNumeric(rLW!Wert) Or rLW!Wert Like "1:*" Then
    If u <> -99 And CDbl(Replace(Replace(rLW!Wert, "1:", ""), ".", ",")) < u Then pKz = "p"
    If o <> -99 And CDbl(Replace(Replace(rLW!Wert, "1:", ""), ".", ",")) > o Then pKz = "p"
   End If
'   End If
   ls = ls + pKz + Chr(13) + Chr(10)
   NB = ""
   If uNG <> "" Or oNG <> "" Then
    If gschl = "w" And (rLW!unw <> "" Or rLW!onw <> "") Then
     NB = nz(rLW!unw,"") + "-" + nz(rLW!onw,"")
    End If
   End If
   If NB = "" And (rLW!unm <> "" Or rLW!onm <> "") Then
    NB = nz(rLW!unm,"") + "-" + nz(rLW!onm,"")
   End If
   ls = ls + NB + Chr(13) + Chr(10)
   If IsNull(rLW!gruppe) Then
    altGruppe = 0
   Else
    altGruppe = Val(rLW!gruppe)
   End If
   rLW.Move 1
 Loop
 
 rDat.Close
 rLW.Close
 LaborString = ls
' Open "c:\le.txt" For Output As #2
' Print #2, LS
' Close #2
 End Function ' LaborString
Function DMPAusgeb(Optional ID&, Optional obAnzeig As Boolean = -1)
 Const ErgebDatei$ = aVerz + "DMP.txt"
 Open ErgebDatei For Output Access Write Lock Write As #1
 DMPAusgeb = DMPString(ID)
 Print #1, DMPAusgeb
 Close #1
 If obAnzeig Then Call Shell("notepad " + ErgebDatei, vbNormalFocus)
End Function

Function TherapieArtEinzelnFestlegen(Pat_ID&)
Dim rAna As DAO.Recordset
Call SysCmd(acSysCmdSetStatus, "Therapiearten festlegen ...")
Set rAna = TabÖff("Anamnesebogen", "Pat_ID")
If rAna.BOF Then Exit Function
rAna.Seek "=", Pat_ID
If Not rAna.NoMatch Then
 rAna.Edit
 rAna!Ther1 = TherArt(Pat_ID, -1, rAna)
 rAna!TherAkt = TherArt(Pat_ID, 0, rAna)
 If rAna!Pat_ID = 0 Then
  rAna.CancelUpdate ' kommt auf mir noch unbekanntem Weg zustande
  Stop
 Else
  rAna.Update
 End If
End If
SysCmd acSysCmdClearStatus
End Function ' TherapieArtEinzelnFestlegen(Pat_id&)
Function TherEintr()
Dim rAna As DAO.Recordset
Call SysCmd(acSysCmdSetStatus, "Therapiearten festlegen ...")
Set rAna = TabÖff("Anamnesebogen", "Auswahl")
If rAna.BOF Then Exit Function
rAna.MoveFirst
Do While Not rAna.EOF
 Call SysCmd(acSysCmdSetStatus, "Therapiearten festlegen für: " + GesNam(rAna))
 rAna.Edit
 rAna!Ther1 = TherArt(rAna!Pat_ID, -1, rAna)
 rAna!TherAkt = TherArt(rAna!Pat_ID, 0, rAna)
 If rAna!Pat_ID = 0 Then
  rAna.CancelUpdate ' kommt auf mir noch unbekanntem Weg zustande
  Stop
 Else
  rAna.Update
 End If
 rAna.Move 1
Loop
SysCmd acSysCmdClearStatus
End Function ' TherEintr()
Function testTherArt$()
 Dim rAna As DAO.Recordset, Pat_ID&
 Pat_ID = 167
 Set rAna = TabÖff("Anamnesebogen", "Pat_id")
 rAna.Seek "=", Pat_ID
 testTherArt = TherArt(Pat_ID, 0, rAna, CDate("17.11.04"))
End Function ' testtherart$()
Function TherArt$(Pat_ID&, Optional obanf%, Optional rAna As DAO.Recordset, Optional VorDat As Date)
Dim insz%, obIns As Boolean, obAnal As Boolean, obGlib As Boolean, obMetf As Boolean, obGlucI As Boolean, obSHGlin As Boolean, obGlit As Boolean
Dim rAnaName$
On Error Resume Next
rAnaName = rAna.Name
If rAnaName = "" Then
  Set rAna = TabÖff("Anamnesebogen", "Pat_id")
  rAna.Seek "=", Pat_ID
  If rAna.NoMatch Then Exit Function
End If
On Error GoTo fehler
 Select Case LCase(rAna!Diabetestyp)
  Case 1, 2, "g", "s", "p"
       Dim rMed As DAO.Recordset, rMA As DAO.Recordset, obRezPumpe As Boolean, KurzMed$, obRezIns As Boolean
       obRezPumpe = False
       obRezIns = False
       If Not obanf Then ' aus den Rezepten über Insulin, Analoga und Pumpe Rückschlüsse ziehen, falls nicht obanf
        sql1 = "SELECT DISTINCT Medikament FROM [" + QMdbAkt + "].RezeptEintraege where pat_id = " & CStr(rAna!Pat_ID)
        If VorDat <> CDate(0) Then sql1 = sql1 + " and zeitpunkt < #" + Format(VorDat + 1, "mm\/dd\/yyyy") + "#"
        Set rMed = Dtb.OpenRecordset(sql1, dbOpenDynaset)
        Set rMA = TabÖff("MedArten", "Medikament")
        If Not rMed.BOF Then
         rMed.MoveFirst
         Do While Not rMed.EOF
          KurzMed = UCase(Left(rMed!Medikament, IIf(InStr(Mid(rMed!Medikament, 4), " ") > 0, InStr(Mid(rMed!Medikament, 4), " ") - 1 + 4, Len(rMed!Medikament))))
          rMA.Seek "=", KurzMed
          If Not rMA.NoMatch Then
           If rMA!Puzu Then
            obRezPumpe = True
            Exit Do
           ElseIf rMA!Ins Or rMA!AnAl Then
            obRezIns = True
           End If
          End If
          rMed.MoveNext
         Loop
        End If
        rMed.Close
       End If
       If InStr(IIf(obanf, rAna!Ther1, rAna!TherAkt), "?") > 0 Then TherArt = ""
       If rAna!Insulinpumpe Or obRezPumpe Then
         TherArt = "CSII"
       Else
         On Error GoTo fehler
         Call TherAuskunft(rAna!Pat_ID, obanf, insz, VorDat, obIns, obAnal, obGlib, obMetf, obGlucI, obSHGlin, obGlit)
         On Error Resume Next
         Select Case insz
          Case 0
           If obRezIns Then
            If IsNull(TherArt) Or TherArt = "" Then TherArt = "I/CT?"
           Else
            If obGlib Or obMetf Or obGlucI Or obSHGlin Or obGlit Then
             TherArt = "OAD"
            Else
             If obIns Or obAnal Then
               MsgBox "Widerspruch in der Funktion TherAuskunft!"
             End If
             If rAna!Diabetestyp = "1" Then
              If IsNull(TherArt) Or TherArt = "" Then TherArt = "(I)CT?"
             Else
              If IsNull(TherArt) Or TherArt = "" Then TherArt = "Diät?"
             End If
            End If
           End If
          Case 1, 2
           If obGlib Or obGlucI Or obSHGlin Or obGlit Then
            TherArt = "Komb"
           Else
            TherArt = "CT"
           End If
          Case Is >= 3
           TherArt = "ICT"
         End Select
       End If
  End Select
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
 Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in TherArt/" + AnwPfad)
  Case vbAbort: Call MsgBox("Höre auf"): End
  Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
  Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
 End Select
End Function ' Therart
Sub GibWerteAus(Optional obAnzeig As Boolean = True)
 do_GibWerteAus obAnzeig
' DiagnosenAusgeb ID, obanzeig
' DMPAusgeb ID, obanzeig
End Sub ' (Optional obAnzeig As Boolean = True)
Sub FormMerk()
 AktID = 0
 On Error Resume Next
 AktID = Forms(Anmnb)!Pat_ID
 On Error GoTo 0
 altRecordSource = ""
End Sub ' FormMerk()
Sub FormAufruf()
 SysCmd acSysCmdSetStatus, "FormAufruf ..."
 If AktID <> 0 Then
  DoCmd.Close acForm, Anmnb, acSaveYes ' "Anamnesebogen nach Name"
  DoCmd.OpenForm Anmnb '"Anamnesebogen nach Name" dtb.Containers(2).Documents(0).Name
  DoCmd.Maximize
  On Error Resume Next
  Forms!Anamnesebogen.Recordset.FindFirst "Pat_ID = " + CStr(AktID) ' Wieder alten Datensatz wählen
  On Error GoTo 0
  If Forms!Anamnesebogen.Recordset.NoMatch() Then
   altRecordSource = Forms!Anamnesebogen.RecordSource
   Forms!Anamnesebogen.RecordSource = "Anamnesebogen alle Datensätze"
   On Error Resume Next
   Forms!Anamnesebogen.Recordset.FindFirst "ID = " + CStr(AktID) ' Wieder alten Datensatz wählen
   On Error GoTo 0
  End If
 End If
End Sub ' FormAufruf()
Sub FormRestoreSource()
Dim i%
If Forms!Anamnesebogen.Name <> Anmnb Then
 DoCmd.Close acForm, Anmnb, acSaveYes
 DoCmd.OpenForm Anmnb
End If
If altRecordSource <> "" Then Forms!Anamnesebogen.RecordSource = altRecordSource
End Sub ' FormRestoreSource()
Sub do_GibWerteAus(Optional obAnzeig As Boolean = True)
Const ErgebDatei$ = aVerz + "Anamnese.txt"
Dim te$
Call FormMerk
Call FormAufruf
te = machWertString(AktID, obAnzeig)
Open ErgebDatei For Output Access Write Lock Write As #1
If te <> "" Then
 Print #1, te
End If
Close #1
Call FormRestoreSource
If obAnzeig Then Call Shell("notepad " + ErgebDatei, vbNormalFocus)
End Sub ' do_GibWerteAus(Optional obAnzeig As Boolean = True)
Function machWertString$(Pat_ID&, Optional obAnzeig As Boolean = False)
Dim rAn As DAO.Recordset, FNr&
Dim te$, Tü$, Add$, Tneu$, Descr$
Dim fld, fldv$
Dim obNZ As Boolean ' neue Zeile
Dim obkomma As Boolean ' Komma
Dim keinName As Boolean
Dim obVorWort As Boolean ' ob Vorwort exisitiert
Dim VorWort$ ' Teil der überschrift für mehrere Tabellenfelder
Dim DPPos% ' Position des Doppelpunktes
Dim bmival#, stelle%
Dim j%, K%
Dim obnadp As Boolean
Dim condens$, rechts$, links$
Set rAn = TabÖff("Anamnesebogen", "Pat_id") 'Forms(Anmnb).RecordsetClone
rAn.Seek "=", Pat_ID
'ran.FindFiran ("Pat_ID = " + CStr(AktID)) 'CStr(ID))
If rAn.NoMatch Then GoTo schluss
te = ""
VorWort = ""
For K = IIf(obAnzeig, 0, 7) To rAn.Fields.Count - 1
 Set fld = rAn.Fields(K)
 Descr = ""
 On Error Resume Next
 Descr = Dtb.TableDefs(fld.SourceTable).Fields(fld.SourceField).Properties("Description")
 FNr = Err.Number
 On Error GoTo 0
 Err.Clear
 If FNr <> 0 Then Descr = Dtb.TableDefs(fld.SourceTable).Fields(fld.SourceField).Name
 
 obNZ = False
 If Left(Descr, 1) = "^" Then
  obNZ = True
  If te <> "" Then
'   te = UCase(Left(te, 1)) + Mid(te, 2)
   If te <> "" Then
'    If Right(te, 1) <> "." Then te = te + "."
'    Print #1, te ' letzte Runde ausgeben
    If Right(te, 1) <> Chr(10) Or te = "" Then te = te + IIf(Right(te, 1) = ".", "", ".") + Chr(13) + Chr(10)
   End If
'   te = ""
  End If
  Descr = LTrim(Mid(Descr, 2))
 End If
 
 obVorWort = False
 If Left(Descr, 1) = "<" Then
  obVorWort = True
  Descr = LTrim(Mid(Descr, 2))
 End If
 
 obkomma = False
 If Left(Descr, 1) = "," Then
  obkomma = True
  Descr = LTrim(Mid(Descr, 2))
 End If
 
 keinName = False
 If Left(Descr, 1) = "-" Then
  keinName = True
  Descr = LTrim(Mid(Descr, 2))
 End If
 
 If Left(Descr, 1) = ":" Then Descr = IIf(keinName, "", fld.Name) + ":"
 DPPos = InStr(Descr, ":")
 If DPPos > 0 Then
  If Not obVorWort Then
   VorWort = Left(Descr, DPPos - 1) + ":"
   Descr = VorWort + " " + Mid(Descr, DPPos + 1)
  End If
 End If
 
 If obVorWort Then
  If Right(te, 1) = Chr(10) Then
   Descr = VorWort + " " + LTrim(Descr)
  End If
 End If
 
 If obkomma Then Descr = ", " + Descr

 If Not IsNull(fld) Then
'  If InStr(fld.name, "chwanger") > 0 Then
'    MsgBox "Achtung: Feld Schwanger: " + fld
'  End If
  If fld <> "False" And fld <> "" Then
   Add = ""
   fldv = Replace(fld.Value, "j ", "ja ")
   If fldv = "j" Then fldv = "ja"
   
   Select Case fld.Name
    Case "Aufschreiben", "Bluthochdruck"
     fldv = Replace(Replace(fldv, "n ", "nein "), "n,", "nein,")
   End Select
   Select Case fld.Name
    Case "Aufschreiben", "Bluthochdruck", "BDselbst", "Schwanger", "Netzhaut gelasert", "Sehminderung unbehebbar", "Diabet Nierenschaden", "Albumin zuletzt", _
         "andere Nierenerkrankung", "Herzkrankheit", "Angina pectoris", "Herzinfarkt", "PTCA oder Stent", "Herzschwäche", _
         "Hirndurchblutungsstörung", "Schlaganfall", "Beindurchblutungsstörung", "Schaufensterkrankheit", "Geschwür", _
         "Amputation", "Ameisenlaufen", "Druckstellen", "Verformungen", "Neue Fußkomplikationen", "Entleerungsstörungen Magen", _
         "Entleerungsstörungen Harnblase", "Schwindel Aufstehen", "Folgeerkrankungen Haut", "Bewegungseinschränkungen", _
         "Sexualstörung", "Alkohol", "Tabak", "Weitere Medikation", "Liphypertrophien Abdomen", "Liphypertrophien Beine", _
         "Liphypertrophien Arme", "Hyperkeratosen", "Ulcera", "Unterzucker pM", "Keto pa", "BZgr300 pM", _
         "Fußpflege", "Podologie", "Einlagen", "BZMessungen pW", "BZMessungen pW ndE", "BZMessungen p W nachts", _
         "BZMessungen Selbst", "Aufschreiben", "BeinödVen" ' "Schulung", "DMP",
     If fldv = "j" Or fldv = "J" Then
      Add = "ja"
     ElseIf fldv = "n" Or fldv = "0" Or fldv = "N" Or fldv = "-" Or fldv = "o.B" Or fldv = "o. B" Then
      Select Case fld.Name
       Case "Unterzucker pM", "Keto pa", "BZgr300 pM", "BZMessungen pW", "BZMessungen pW ndE", "BZMessungen p W nachts"
        Add = "keine"
       Case Else
        Add = "nein"
      End Select
     ElseIf fldv = "u" Then
      Add = "unbekannt"
     Else
      If Left(fldv, 2) = "n " Then
       Add = Replace(fldv, "n ", "nein ", , 1)
      Else
       Add = CStr(fldv)
      End If
     End If
     fldv = Add
    Case "ASR", "PSR", "Puls Leiste", "Puls Kniekehle", "Puls Atp", "Puls Adp"
     Add = PulsParse(fldv)
     If InStr(Add, "bi") <> 0 Or InStr(Add, "mono") <> 0 Or InStr(Add, "post") <> 0 Then
      Descr = Replace(Descr, "Pulse", "Pulse/Dopplersignale")
     End If
     fldv = Add
   End Select
'   Select Case fld.Name
'    Case "Oberflächensensibilität", "Monofilamenttest", "Kalt-Warm", "Vibration IK", "Vibration Großzehe"
'     j = InStr(fldV, "/")
'     If j > 0 And InStr(j + 1, fldV, "/") = 0 Then
'      fldV = Trim(fldV) + ", " + Trim(fldV)
'     End If
'   End Select
   Select Case fld.Name
    Case "Liphypertrophien Abdomen", "Liphypertrophien Beine", "Liphypertrophien Arme", "Beinbefund", "Hyperkeratosen", _
         "Ulcera", "Kraft Zehenheber", "Kraft Zehenbeuger", "Kraft Knie", "ASR", "PSR", "Oberflächensensibilität", _
         "Puls Leiste", "Puls Kniekehle", _
         "Puls Atp", "Puls Adp", "Blutdruckwerte"
     If InStr(fldv, ",") > 0 Then
      fldv = "rechts: " + LTrim(Mid(fldv, 1, InStr(fldv, ","))) + " links: " + LTrim(Mid(fldv, InStr(fldv, ",") + 1))
     End If
   End Select
   
   Select Case fld.Name
    Case "Blutdruckwerte", "RR"
     fldv = Replace(fldv, "arr.", "arrhythmisch")
     fldv = Replace(fldv, "arrh.", "arrhythmisch")
     If InStr(fldv, "arrh") <> InStr(fldv, "arrhythmisch") Then
      fldv = Replace(fldv, "arrh", "arrhythmisch")
     ElseIf InStr(fldv, "arr") <> InStr(fldv, "arrhythmisch") Then
      fldv = Replace(fldv, "arr", "arrhythmisch")
     End If
    Case "Herz", "Carotiden"
     fldv = Replace(fldv, "SG", "Strömungsgeräusch")
   End Select
   obnadp = False
   j = InStr(LCase(fldv), "nadp")
   If j > 0 Then
    obnadp = True
    fldv = Left(fldv, j - 1) + LTrim(Right(fldv, Len(fldv) - j - 3))
   End If
   
   Select Case fld.Name
    Case "Tkz", "Pat_ID", "HANr", "HANr2", "letzte Änderung", "Diagnosen", "Vorgestellt", "Schulung", "DMP", "Vorname", "NVorsatz", "Titel", "Anrede", "RRTurboMed", "Versicherung", "DMSchulz", "DMSchl", "RRSchulz", "DMPhier", "AktZeit", "Ther1", "TherAkt", "Prim", "obAn1eing", "obAn2eing", "obAnAeing", "obCheck", "obBZausgew", "obOSaufgek", "obPodAufgek", "obMBlAusgeh", "obSchulaufgek", "obDMPaufgekl", "obMedNetz", "Versicherungsart", "Hausarzt", "ob"
    Case "Diabetestyp"
     Add = " " + DTyp(rAn.Fields(fld.Name))
    Case "NachName"
     Add = CStr(fldv) + ", " + rAn!Vorname
     If Not IsNull(rAn!Titel) Then
      If Len(rAn!Titel) > 0 Then
       Add = rAn!Titel + " " + Add
      End If
     End If
     If Not IsNull(rAn!Anrede) Then
      If Len(rAn!Anrede) > 0 Then
       Add = rAn!Anrede + " " + Add
      End If
     End If
    Case "Diabetes seit"
     If fldv = "bu" Or fldv = "b" Then
      Descr = ", "
      Add = "bisher unbekannt"
     Else
      Add = CStr(fldv)
     End If
    Case "Weitere Anamnese"
     Add = CStr(fldv)
    Case "Tendenz"
     Select Case fldv
      Case "a"
       Add = "abnehmend"
      Case "g"
       Add = "gleichbleibend"
      Case "z"
       Add = "zunehmend"
     End Select
    Case "Tabletten seit", "Insulin seit"
     If fldv = "-" Then Add = "" Else Add = CStr(fldv)
    Case "Insulinpumpe"
     If fldv Then Add = "ja"
    Case "Größe", "Gewicht"
     If fldv = "0" Then Add = "" Else Add = CStr(fldv)
     If fld.Name = "Gewicht" Then
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
       Add = Add + ", BMI = " + Format(bmival, "###.#") + " kg/m²"
      End If
      On Error GoTo 0
     Else ' fld.name = "Größe"
      Add = Add + " cm"
     End If
    Case "Essenszeit früh", "Essenszeit mittags", "Essenszeit abends", "Essenszeit spät"
     If fldv = "-" Then
      Add = "-"
     Else
      Add = CStr(fldv) + " Uhr"
     End If
    Case "Spritz-Eß-Abstand früh", "Spritz-Eß-Abstand mittags", "Spritz-Eß-Abstand abends"
     Add = CStr(fldv) + " min"
    Case "Spritzstelle früh", "Spritzstelle mittags", "Spritzstelle abends", "Spritzstelle nachts"
     Add = Spritzstelle(fldv)
'    Case "Puls Leiste"
'     Add = "Leiste " + CStr(fldV)
'    Case "Liphypertrophien Abdomen"
'     Add = "Abdomen " + CStr(fldV)
    Case "UZ rechtzeitig"
     Select Case fldv
      Case "i"
       Add = "immer"
      Case "m"
       Add = "meist"
      Case "n"
       Add = "nie"
      Case Else
       Add = Replace(Replace(Replace(fldv, "j ", "ja "), "m ", "meist "), "n ", "nie ")
       If Left(Add, 2) = "i " Then Add = Replace(Add, "i ", "immer ", , 1)
     End Select
    Case "letztes HbA1c", "vorherige Werte"
     Add = fldv + " %"
     If InStr("0123456789", Right(Trim(fldv), 1)) > 0 And Trim(fldv) <> "" Then
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
     On Error GoTo 0
    Case "fremde Hilfe pa", "Bewußtlos pa"
     If fldv = "0" Or fldv = "n" Or fldv = "-" Then
       Add = "kein Mal"
     ElseIf Left(fldv, 1) = "(" Then
      Add = fldv
     Else
      Add = fldv + " mal"
     End If
    Case "Blutdruckwerte", "RR"
'     If InStr("0123456789", Right(Trim(fldV), 1)) > 0 Then
'      Add = CStr(fldV) + " mm Hg"
'     Else
       Add = CStr(fldv)
'     End If
     stelle = InStr(Add, "/")
'     If stelle = 0 Then
'      If InStr(fldV, "mm Hg") = 0 Then
'       Add = CStr(fldV) + " mm Hg"
'      End If
'     End If
     While stelle <> 0
      stelle = stelle + 1
      While InStr("0123456789", Mid(Add, stelle, 1)) > 0 And stelle <= Len(Add)
       stelle = stelle + 1
      Wend
      If Mid(Add, stelle, 2) <> "mm" And Mid(Add, stelle + 1, 2) <> "mm" Then
       Add = Left(Add, stelle - 1) + " mm Hg" + Right(Add, Len(Add) - stelle + 1)
      End If
      stelle = InStr(stelle + 1, Add, "/")
     Wend
     If Trim(LCase(Add)) = "n" Then Add = "normal"
    Case "Tabak"
     If InStr("0123456789", Right(Trim(fldv), 1)) > 0 And Not IsNumeric(Right(Trim(fldv), 4)) And Trim(fldv) <> "" Then
      Add = CStr(fldv) + " Zig./d"
     Else
      Add = CStr(fldv)
     End If
    Case "BZWerte v d Essen", "BZWerte n d Essen"
     If InStr("0123456789", Right(Trim(fldv), 1)) > 0 And Trim(fldv) <> "" Then
      Add = CStr(fldv) + " mg/dl"
     Else
      Add = CStr(fldv)
     End If
    Case "Jahr letzte Diabetesschulung"
     If fldv = "-" Then
      Add = "keine"
     Else
      Add = CStr(fldv)
     End If
    Case "Ort Schulung"
     If fldv = "-" And rAn![Jahr letzte Diabetesschulung] = "-" Then
      fldv = ""
     Else
      Select Case fldv
       Case "KD"
        Add = "Klinikum Dachau"
       Case "KMB"
        Add = "Klinikum Bogenhausen"
       Case "KMS"
        Add = "Klinikum Schwabing"
       Case Else
        Add = CStr(fldv)
      End Select
      If InStr(Add, "linikum") > 0 Or InStr(Add, "rankenh") > 0 Then
       Descr = "im"
      ElseIf InStr(Add, "Dr.") > 0 Then
       Descr = "bei"
      Else
       Descr = "in"
      End If
     End If
    Case "Augensp Befund", "Herz", "Lunge", "Bauch", "Mundhöhle", "Carotiden"
     If LCase(fldv) = "n" Or LCase(fldv) = "ob" Or LCase(fldv) = "ok" Or LCase(Left(fldv, 3)) = "o.b" Then
      Select Case fld.Name
       Case "Carotiden"
        Add = "kein Strömungsgeräusch"
       Case Else
        Add = "unauffällig"
      End Select
     ElseIf fldv = "neg" Then
      Add = "negativ"
     Else: Add = CStr(fldv)
     End If
    Case "erhöht?"
     If Trim(fldv) <> "" Then
      If InStr("n-nein", LCase(fldv)) = 0 Then
       Add = "normal"
      ElseIf InStr("j+ja", LCase(fldv)) = 0 Then
       Add = "erhöht"
      End If
     End If
    Case "Dialyse", "Bypass kardial", "Bypaß peripher"
     If LCase(fldv) = "falsch" Then
      Add = ""
     ElseIf LCase(fldv) = "wahr" Then
      Add = "ja"
     End If
    Case "Herzinfarkt wann", "Verformungen Beschreibung"
     If fldv = "n" Or fldv = "-" Then
      Add = ""
     Else
      Add = CStr(fldv)
     End If
    Case "Kraft Zehenheber", "Kraft Zehenbeuger", "Kraft Knie"
     If fldv = "n" Then
      Add = "seitengleich normal"
     ElseIf fldv = "sg" Or fldv = "sehr gut" Or fldv = "sehr kräftig" Then
      Add = "seitengleich sehr kräftig"
     Else
      Add = CStr(fldv)
     End If
    Case "Oberflächensensibilität"
     If fldv = "n" Then
      Add = "seitengleich ungestört"
     ElseIf fldv = "?" Then
      Add = "?/?"
     Else
      Add = CStr(fldv)
     End If
    Case "Monofilamenttest", "Kalt-Warm", "Vibration IK", "Vibration Großzehe"
      Call Bruchteile(fldv, rechts, links)
      If IsNumeric(Right(rechts, 1)) Then
       If InStr(rechts, "/") = 0 Then
        Select Case fld.Name
         Case "Monofilanmenttest", "Kalt-Warm"
          rechts = rechts + "/5"
         Case Else
          rechts = rechts + "/8"
        End Select
       End If
      End If
      If IsNumeric(Right(links, 1)) Then
       If InStr(links, "/") = 0 Then
        Select Case fld.Name
         Case "Monofilanmenttest", "Kalt-Warm"
          links = links + "/5"
         Case Else
          links = links + "/8"
        End Select
       End If
      End If
      If Replace(Replace(Replace(rechts, "rechts", ""), "re ", ""), "re.", "") = Replace(Replace(Replace(links, "links", ""), "li ", ""), "li.", "") Then
        Add = IIf(InStr(rechts, "bds"), "", "bds. ") + Replace(Replace(Replace(rechts, "rechts", ""), "re ", " "), "re.", "")
      Else
       If InStr(rechts, "rechts") = 0 And InStr(rechts, "re ") = 0 And InStr(rechts, "re.") = 0 Then
        rechts = "re: " + rechts
       End If
       If InStr(links, "links") = 0 And InStr(links, "li ") = 0 And InStr(links, "li.") = 0 Then
        links = "li: " + links
       End If
       Add = rechts + ", " + links
      End If
    Case "WS", "NL", "NNH"
     If fldv = "n" Then
      Add = "kein Klopfschmerz"
     Else: Add = Replace(fldv, "KS", "Klopfschmerz")
     End If
    Case "Zähne"
     Add = Replace(Replace(fldv, "TP", "Teilprothese"), "VP", "Vollprothese")
    Case "SD"
     If fldv = "n" Then
      Add = "nicht vergrößert tastbar"
     Else: Add = CStr(fldv)
     End If
    Case "Neuro sonst"
     Add = Replace(Replace(fldv, "KR", "Konvergenzreaktion"), "LR", "Lichtreaktion")
    Case Else
     Add = CStr(fldv)
   End Select
   If obnadp Then Add = "nach Angabe d.Pat. " + Add
   If obNZ And Add = "" And fld.Name <> "Insulinpumpe" Then Add = "-"
' 2. Zeile Ende gestrichen: 'IIf(Descr = ":", fld.Name, "") +
   If Add <> "" Then _
    te = te + IIf(Right(te, 1) = Chr(10) Or te = "", "", IIf(Left(Descr, 1) = ",", "", " ")) + _
         Trim(IIf(Right(te, 1) = Chr(10) Or te = "", IIf(Left(Descr, 2) = ", ", Mid(Descr, 3), Descr), Descr)) + IIf(te = "", "", " ") + Add
  End If ' fld <> "False" And fld <> "" Then
 End If ' not isnull(fld)
Next
te = UCase(Left(te, 1)) + Mid(te, 2)
If te <> "" Then
 If Right(te, 3) <> "." + Chr(13) + Chr(10) Then te = Left(te, Len(te) - 2) + "." + Chr(13) + Chr(10)
End If
machWertString = te
schluss:
rAn.Close
End Function 'machWertString
Function Bruchteile(fldv$, re$, li$)
      Dim tp%  ' Trennposition
      Dim tp1% ' Trennposition 1
      tp = InStr(fldv, "|")
      If InStr(fldv, ",") > 0 And InStr(fldv, "bds") = 0 Then
       re = Left(fldv, InStr(fldv, ",") - 1)
       li = Right(fldv, Len(fldv) - InStr(fldv, ","))
      ElseIf tp > 0 Then
       re = Left(fldv, tp - 1)
       li = Mid(fldv, tp + 1)
       tp = InStr(re, "/")
       If tp > 0 Then re = Left(re, tp - 1)
       tp = InStr(li, "/")
       If tp > 0 Then li = Left(li, tp - 1)
      Else
       re = fldv
       li = fldv
      End If
      re = LTrim(Trim(re))
      li = LTrim(Trim(li))
End Function ' Bruchteile(fldv$, re$, li$)
Function Spritzstelle(Abkü)
  Select Case LCase(CStr(Abkü))
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
    Spritzstelle = "meist Bauch"
   Case "meist os"
    Spritzstelle = "meist Oberschenkel"
   Case "b/os", "os/b"
    Spritzstelle = "wechselnd Bauch/Oberschenkel"
   Case "b/o/a"
    Spritzstelle = "wechselnd Bauch/Oberschenkel/Arm"
   Case "b os"
    Spritzstelle = "Bauch bzw. Oberschenkel"
   Case "a/b", "b/a"
    Spritzstelle = "wechselnd Arm/Bauch"
  End Select
End Function ' Spritzstelle(Abkü)

