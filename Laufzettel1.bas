Attribute VB_Name = "Laufzettel"
Option Explicit
 Const Verzeich$ = "p:\plz"

Function doPatientenlaufzettel()
' für PDF lesen
 Dim PDTextS As Acrobat.CAcroPDTextSelect
 Dim Result&, PDDoc As Acrobat.CAcroPDDoc
 Dim PDPage As Acrobat.CAcroPDPage
 Dim PDHili As Acrobat.CAcroHiliteList
 Dim B As Boolean, erg$
 Dim PDFStream$, Filename$
' Strings Falls Text von Anfang ausgelesen werden soll (bzw bis zum Ende)
 Dim FromFirst$, ToLast$
 Dim Str$, NTL&
 Dim i%, j&, splitt$(), s2$(), pid$, Datum As Date
 Dim Uhrzeit$
 Filename = Verzeich & "\TMFTools.pdf"
' Filename = "p:\Hausmann Karteikarte.pdf"
 FromFirst = Chr(169) & Chr(170) & Chr(172)
 ToLast = Chr(163) & Chr(165) & Chr(164)
 Set PDDoc = CreateObject("AcroExch.pdDoc")
 Result = PDDoc.Open(Filename)
 If Not Result Then
  MsgBox "Fehler beim Öffnen von: " & Filename, vbCritical, "Patientenlaufzettel:"
  Exit Function
 End If
' Nehme die Erste Seite - Index 0
 Do
  Set PDPage = PDDoc.AcquirePage(j)
  ' Erzeuge ein Highlight Objekt und weise ihm 2000 Elemente bei (keine Grenzprobleme)
  Set PDHili = CreateObject("AcroExch.HiliteList")
  B = PDHili.Add(0, 32000)
  ' Erzeuge eine Textauswahl aus dem gesamten Text
  On Error Resume Next
  Set PDTextS = PDPage.CreatePageHilite(PDHili)
  If Err.Number <> 0 Then Exit Do
  On Error GoTo fehler
  ' Hole Anzahl der "Textblöcke"
  NTL = PDTextS.GetNumText
  ' Gebe den Text der Textauswahl zurück
  For i = 0 To NTL - 1
   Str = Str & PDTextS.GetText(i)
  Next i
  Str = Str & vbCrLf
  j = j + 1
 Loop
 Result = PDDoc.Close
 Set PDPage = Nothing
 Set PDHili = Nothing
 Set PDTextS = Nothing
 Set PDDoc = Nothing
 splitt = Split(Str, vbCrLf)
 s2 = Split(splitt(0), " ")
 For j = 0 To UBound(s2)
  If InStrB(s2(j), "-") > 0 Then
   s2 = Split(s2(j), "-")
   Datum = CDate(s2(0))
   Exit For
  End If
 Next j
 erg = Dir(Verzeich & "\*")
 Do While erg <> ""
  If InStrB(erg, "Patientenlaufzettel") <> 0 Then
   Kill Verzeich & "\" & erg
  End If
  erg = Dir
 Loop
 DoEvents
 For j = 0 To UBound(splitt)
  If InStrB(splitt(j), " ") <> 0 Then
   i = InStr(splitt(j), " ")
   pid = Left$(splitt(j), i - 1)
   If IsNumeric(pid) Then
    If InStrB(pid, ".") = 0 And InStrB(pid, ",") = 0 And InStrB(pid, ";") = 0 Then
     Uhrzeit = ""
     i = InStr(splitt(j), ":")
     If i <> 0 Then
      Uhrzeit = Mid(splitt(j), i - 2, 5)
     End If
     Call dodoPLZ(CLng(pid), Datum, Uhrzeit)
    End If
   End If
  End If
 Next j
 MsgBox "Fertig!"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doPatientenlaufzettel/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function

Function getDiast$(Wert$)
 Dim psl&, psp&
 psl = InStr(Wert, "/")
 If psl <> 0 Then
  psp = InStr(psl + 2, Wert, " ") - 1
  If psp = -1 Then psp = Len(Wert)
 End If
 getDiast = Trim$(Mid$(Wert, psl + 1, psp - psl))
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getDiast/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getDiast
Function dodoPLZ(pid$, Optional Datum As Date, Optional Uhrzeit$) ' Patientenlaufzettel
 Dim rNam As New ADODB.Recordset, rAn As New ADODB.Recordset, rFl As New ADODB.Recordset, rFlb As New ADODB.Recordset, FlbZ&
 Dim rDia As New ADODB.Recordset, rsMB As New ADODB.Recordset
 Dim lsql$, lMP$, i&, Größe!, bmi!, bmiS$, Gewicht!, obpath%, Wert$, WertNum!, j&, k&, ia&, ie&
 Const zl% = 30     ' Zahl der Parameter
 Const SchulPID& = 1789
 Dim RRi% ' Parameter für Blutdruck
 Dim Neuri% ' Parameter für Neurostatus
 Dim Albi% ' Parameter für Albumin
 Dim Gwi% ' Paraemter für Gewicht
 Dim RRVgli% ' Parameter für RR-Vergleich
 Dim Urini% ' Parameter für Urin
 Dim Kreai%, GFRi% ' Parameter für Kreatinin / GFR
 Dim Schuli% ' Parameter für Schulungen
 Dim SchulEintri% ' Parameter für Schulungseinträge
 Dim Hbi%, Ferri%, B12i%, Foli% ' Parameter für Hb, Ferritin, Vit B12, Folsre
 Dim Leuki%, CRPi%, TSHi%, fT4i%, fT3i%, TGi%, GGTi%, GPTi%
 Dim obNP%, pathAlbZ%, AlbZ%, pathKreaZ%, KreaZ% ' ob Nephropathie vorliegt
 Dim Caroi%, Caroai% ' Carotiduplex durchgeführt, alle
 Const zmaxges% = 8 ' Maximale Zeilenzahl pro Parameter
 Dim DateiName$
 Dim obMB% ' ob Merkblatt Fußsyndrom mitgegeben
 Dim obAvan% ' ob Avandia verordnet wird
 Dim rs(zl) As New ADODB.Recordset, rszahl As New ADODB.Recordset
 Dim sql$(zl)
 Dim rsz&(zl) ' vorhandene Zahl
 Dim zmax%(zl), rszmax& ' maximal dargestellte Zahl, Maximum davon
 Dim Üs$(zl), Titel$(zl)  ' Überschrift
 Dim Fqmin(zl) ' Mindest-Untersuchungs-Frequenz
 Dim UGrenze!(zl), Grenze!(zl) ' Grenze für pathologische Werte
 Const GrenzeDiast! = 81
 Dim nurpath%(zl) ' 1 = Anzeige nur bei Fehlen eines Wertes und bei pathologischem Wert,
                  ' 2 = Anzeige nur bei pathologischem Wert
                  ' 3 = Anzeige bei Vorhandensein eines Wertes
 Dim loopct%, pathol%(zl, zmaxges) ' pathol(zl,0)= ' ob irgend einer der Werte pathologisch ist
 Dim pathz%(zl), npathz%(zl) ' Zahl der nicht/ Pathologischen
 Dim üw12erg$()
 
 On Error GoTo fehler
 If DBCn Is Nothing Or DBCn.ConnectionString = "" Then Call acon(quelleT)
 rNam.Open "select * from namen where pat_id = " & pid, DBCn, adOpenStatic, adLockReadOnly
 rFl.Open "select k.*, f.* from faelle f join aktfv a using (fid) left join kassenliste k on(((f.VKNr = k.VK) and (f.IK = k.IK))) where f.pat_id = " & pid & " group by f.pat_id", DBCn, adOpenStatic, adLockReadOnly
 If Not rNam.EOF Then
  rAn.Open "select * from anamnesebogen where pat_id = " & pid, DBCn, adOpenStatic, adLockReadOnly
  If Datum = 0 Then Datum = Now
  DateiName = Verzeich & "\" & rNam!Nachname & " " & rNam!Vorname & ",   Pid " & rNam!Pat_id & ", Patientenlaufzettel.html" ' Replace(CDate(Uhrzeit), ":", ".") & " " & Format(Datum, "dd.mm.yyyy") &
  Open DateiName For Output As #355
  Print #355, "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>"
  Print #355, "<HTML>"
  Print #355, "<HEAD>"
  Print #355, "    <META HTTP-EQUIV='CONTENT-TYPE' CONTENT='text/html; charset=windows-1252'>"
  Print #355, "    <TITLE>" & rNam!Nachname & " " & rNam!Vorname & " (Pat_id:" & pid & "): DMP-Zettel vom " & Now & "</TITLE>"
  Print #355, "    <META NAME='AUTHOR' CONTENT='Gerald Schade'>"
  Print #355, "    <META NAME='CREATED' CONTENT=" & Format(Now, "yyyymmdd") & ";" & Format(Now, "hhmmss") & "00'>"
  Print #355, "    <META NAME='CHANGEDBY' CONTENT='Gerald Schade'>"
  Print #355, "    <META NAME='CHANGED' CONTENT=" & Format(Now, "yyyymmdd") & ";" & Format(Now, "hhmmss") & "00'>"
  Print #355, "    <META NAME='CHANGEDBY' CONTENT='Gerald Schade'>"
  Print #355, "    <STYLE TYPE='text/css'>"
  Print #355, "    <!--"
  Print #355, "        @page { size: 21cm 29.7cm; margin: 1cm }" ' text-transform:capitalize
  Print #355, "    @media screen {"
  Print #355, "      #cave { font-weight:bold;color:red;text-decoration:blink; }"
  Print #355, "      #treat { font-weight:bold;color:blue;text-decoration:underline; }"
  Print #355, "      #path { font-weight:bold;color:blue;background-color:yellow;font-size:2.8mm;}"
  Print #355, "    }"
  Print #355, "    @media print {"
  Print #355, "      #cave { font-weight:bold;color:red;text-transform:capitalize;text-decoration:underline;background-color:yellow }"
  Print #355, "      #treat { font-weight:bold;color:blue;background-color:yellow }"
  Print #355, "      #path { font-weight:bold;color:blue;font-size:2.8mm;background-color:yellow }"
  Print #355, "    }"
  Print #355, "      #abstand { font-size:2.5mm }"
  Print #355, "      #unauff { font-weight:normal;font-size:2.8mm }"
  Print #355, "      #schmal { width:2.5em;border-left-width:1.5mm }"
  Print #355, "      body { margin:0 }"
  Print #355, "      h1 { margin-bottom:0.1cm;font-size:4.5mm;font:normal }"
  Print #355, "      h2 { margin-bottom:0.1cm;font-size:4.5mm;font:normal }"
  Print #355, "      th { border-style:none solid;overflow:auto;font:bold;font-size:3mm }" ' padding-left:0.5em;
  Print #355, "      td { border-style:solid;max-width:15em;border-width:0.1mm;margin:0 }"
  Print #355, "      td p { text-align:left;overflow:auto;max-height:10mm;}"
  Print #355, "      tr { vertical-align:top }"
  Print #355, "      table { font-size:2.8mm; border-collapse:collapse; }"
  Print #355, "    -->"
  Print #355, "    </STYLE>"
  Print #355, "</HEAD>"
  Print #355, "<BODY LANG='de-DE' DIR='LTR'>"
  
  
  For i = 0 To zl: UGrenze(i) = -1: Grenze(i) = -1: Next i
'  zmax(0) = 6: zmax(1) = 6: zmax(2) = 4: zmax(3) = 4: zmax(4) = 6: zmax(5) = 4: zmax(6) = 8
  For i = 0 To zl: zmax(i) = zmaxges: Next i '8
  lsql = "select * from labor1 where pat_id = " & pid & " union select * from labor2 where pat_id = " & pid
  lMP = "select mp.* from (select pat_id, mpnr from (select mpi.pat_id, mpi.mpnr, mpi.zeitpunkt from (select pat_id, max(zeitpunkt) as zp from medplan mp where pat_id = " & pid & " and zeitpunkt <= now() group by pat_id) as mpl left join medplan mpi on mpl.pat_id = mpi.pat_id and mpl.zp = mpi.zeitpunkt group by mpl.pat_id, mpl.zp, mpnr order by  mpl.pat_id, mpl.zp, mpnr desc) as mpii group by pat_id, zeitpunkt) as innen left join medplan mp on innen.pat_id = mp.pat_id and innen.mpnr = mp.mpnr"
  i = 0
  
  Üs(i) = "Gw"
  Titel(i) = "Körpergewicht [kg]"
  Gwi = i
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where art like 'gewicht%' and pat_id = " & pid & " order by zp desc"
  Fqmin(i) = 4
  i = i + 1
  
  Üs(i) = "Taille"
  Titel(i) = "Taille (Mitte zwischen Hüfte und Beckenkamm) [cm], NB &#x2640;(-80)-88, &#x2642;(-94)-102"
  Gwi = i
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where art like 'taile%' and pat_id = " & pid & " order by zp desc"
  Fqmin(i) = 1
  If rNam!Geschlecht = "m" Then
   Grenze(i) = 102
  Else
   Grenze(i) = 88
  End If
  i = i + 1
  
  Üs(i) = "RR"
  Titel(i) = "Blutdruck [mm Hg]<br>Ziel bei Nephropathie niedriger"
  RRi = i
  sql(i) = "select zeitpunkt zp, rr wert from rr where pat_id = " & pid & " order by zeitpunkt desc"
  Fqmin(i) = 180
  Grenze(i) = 130
  i = i + 1
  
  Üs(i) = "HbA1c"
  Titel(i) = "HbA1c [%], nach DCCT, NB [4,1-6,2]"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü like 'HBA1C%' order by zp desc"
  Fqmin(i) = 4
  Grenze(i) = 6.5
  i = i + 1
  
  Üs(i) = "NeuSt"
  Titel(i) = "`Neuro-Status`, Fußuntersuchung"
  Neuri = i
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where art like 'usdm%' and pat_id = " & pid & " order by zp desc"
  Fqmin(i) = 1
  i = i + 1
  
  Üs(i) = "Chol"
  Titel(i) = "Gesamtcholesterin [mg/dl]"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü in ('CHOL') order by zp desc"
  Fqmin(i) = 1
  Grenze(i) = 200
  i = i + 1
  
  Üs(i) = "HDL"
  Titel(i) = "HDL-Cholesterin [mg/dl]"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü in ('HDL','HDLC') order by zp desc"
  Fqmin(i) = 1
  UGrenze(i) = 40
  i = i + 1
  
  Üs(i) = "LDL"
  Titel(i) = "LDL-Cholesterin [mg/dl]"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü in ('LDL','LDLH','LDLC','LDLH01') order by zp desc"
  Fqmin(i) = 1
  Grenze(i) = 100
  i = i + 1
  
  Üs(i) = "TG"
  Titel(i) = "Triglyceride [mg/dl]"
  TGi = i
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü in ('TRI','TRIG','NTFE') order by zp desc"
  Fqmin(i) = 1
  Grenze(i) = 150
  i = i + 1
  
  Üs(i) = "GGT"
  Titel(i) = "Gamma-GT [U/l], NB &#x2640;-39, &#x2642;-66"
  GGTi = i
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü like '%GT%' order by zp desc"
  Fqmin(i) = 1
  If rNam!Geschlecht = "m" Then
   Grenze(i) = 66
  Else
   Grenze(i) = 39
  End If
  i = i + 1
  
  Üs(i) = "GPT"
  Titel(i) = "GPT (ALAT) [U/l], NB &#x2640;-35, &#x2642;-50"
  GPTi = i
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü like 'GPT%' order by zp desc"
  Fqmin(i) = 1
  If rNam!Geschlecht = "m" Then
   Grenze(i) = 50
  Else
   Grenze(i) = 35
  End If
  i = i + 1
  
  Üs(i) = "Urin"
  Titel(i) = "Urinbefund im Teststreifen"
  Urini = i
  sql(i) = "select * from (select zeitpunkt zp, inhalt wert from eintraege where pat_id = " & pid & " and art like 'urin%') u1 union select * from (select eingang zp, keimzahl wert from laborxbakt natural join laborxus where pat_id = " & pid & " and keimzahl<> '') u2 order by zp desc"
  Fqmin(i) = 1
  i = i + 1
  
  Üs(i) = "Alb/U"
  Titel(i) = "Albuminuie [mg/g Kreatinin], NB < 20 mg/gCrea, bei Nichterrechnen Albumin [mg/l]"
  Albi = i
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü in ('ALBCRE','ALBKRE','ALBQ','ALBU') and (abkü <> 'ALBU' or wert like '%<%') order by zp desc"
  Fqmin(i) = 1
  Grenze(i) = 20
  i = i + 1
  
  Üs(i) = "Krea"
  Titel(i) = "Kreatinin im Serum [mg/dl], Jaffé-Methode, NB &#x2640; -1,1 mg/dl, &#x2642; -1,4 mg/dl"
  Kreai = i
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü in ('KREA', 'KREA02', 'KRE02', 'CREAT') order by zp desc"
  Fqmin(i) = 2
  If rNam!Geschlecht = "m" Then
   Grenze(i) = 1.4
  Else
   Grenze(i) = 1.1
  End If
  i = i + 1
  
  Üs(i) = "GFR"
  Titel(i) = "Glomeruläre Filtrationsrate berechnet [ml/min], NB > 60 mg/dl"
  GFRi = i
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü in ('GFRT', 'CREACL') order by zp desc"
  Fqmin(i) = 0
  UGrenze(i) = 60
  nurpath(i) = 2
  i = i + 1
  
  Üs(i) = "Augen-US"
  Titel(i) = "Augenärztliche Untersuchung"
  sql(i) = "select * from (select zeitpunkt zp, inhalt wert from eintraege where pat_id = " & pid & " and (" + artspez + " and ((inhalt like '%augenb%' and not inhalt like '%augenbl%' and not inhalt like '%augen') or (inhalt like '%augenarzt%' or inhalt like '%augenärzt%') or (inhalt like '% aa%' and not inhalt like '% aag%'))) or (art = 'aa' or art = 'augen'))" & " and zeitpunkt > " & datform(Now() - 550) & ") u1 " & _
     "union select * from (select zeitpunkt zp, dokname wert from dokumente where pat_id = " & pid & " and dokname like '%augen%' and zeitpunkt > " & datform(Now() - 550) & ") u2 order by zp desc"
  Fqmin(i) = 1
  i = i + 1
  
  Üs(i) = "RR-Vgl"
  Titel(i) = "Vergleich des Blutdruckgerätes mit unserem Oberarmmanschettengerät"
  RRVgli = i
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where art like 'rrvgl%' and pat_id = " & pid & " order by zp desc"
  Fqmin(i) = 1
  i = i + 1
  
  Üs(i) = "BZ-Vgl"
  Titel(i) = "Vergleich des Blutzuckergerätes mit Naßchemiegerät"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where art like 'bzvgl%' and pat_id = " & pid & " order by zp desc"
  Fqmin(i) = 1
  i = i + 1
  
  Üs(i) = "Schul"
  Schuli = i
  Titel(i) = "(Gruppen-)Schulungen, Nachschulungen"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where art like 'schul%' and pat_id = " & pid & " order by zp desc"
  Fqmin(i) = 1
  i = i + 1
  
  Üs(i) = "TSH"
  Titel(i) = "TSH [µU/ml], NB 0,27 - 4,2"
  TSHi = i
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü in ('TSH', 'TS1E01', 'TSH-L_K') order by zp desc"
  Fqmin(i) = 2
  UGrenze(i) = 0.27
  Grenze(i) = 4.2
  nurpath(i) = 1
  i = i + 1
  
  Üs(i) = "fT4"
  Titel(i) = "fT4 [ng/dl], NB 0,8-1,8"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü like 'FT4%' order by zp desc"
  fT4i = i
  Fqmin(i) = 0
  UGrenze(i) = 0.8
  Grenze(i) = 1.8
  nurpath(i) = 2
  i = i + 1
  
  Üs(i) = "fT3"
  Titel(i) = "fT3 [pg/ml], NB 2,0 - 4,2"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü like 'FT3%' order by zp desc"
  fT3i = i
  Fqmin(i) = 0
  UGrenze(i) = 2#
  Grenze(i) = 4.2
  nurpath(i) = 2
  i = i + 1
  
  Üs(i) = "Schuleintr"
  Titel(i) = "Schulungseintrag mit Nachname des Patieten"
  SchulEintri = i
  sql(i) = "select zeitpunkt zp, concat (art, ' <- ', inhalt) wert from eintraege where pat_id = 1789 and inhalt like '%" & rNam!Nachname & "%' order by zp desc"
  nurpath(i) = 3
  i = i + 1
  
  Üs(i) = "Hb"
  Titel(i) = "Hämoglobin [g/dl], NB &#x2640;12-16, &#x2642;14-18"
  Hbi = i
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü in ('Hb') order by zp desc"
  Fqmin(i) = 2
  If rNam!Geschlecht = "m" Then
   UGrenze(i) = 14
   Grenze(i) = 18
  Else
   UGrenze(i) = 12
   Grenze(i) = 16
  End If
  nurpath(i) = 1
  i = i + 1
  
  Üs(i) = "Ferritin"
  Ferri = i
  Titel(i) = "Ferritin [ng/ml], NB [15 - 150]"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü like 'FERR%' order by zp desc"
  Fqmin(i) = 0
  UGrenze(i) = 15
  Grenze(i) = 150
  nurpath(i) = 2
  i = i + 1
  
  Üs(i) = "Vit B12"
  B12i = i
  Titel(i) = "Vitamin B12 [pg/ml], NB [211-911]"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü in ('B12','VI1201') order by zp desc"
  Fqmin(i) = 0
  UGrenze(i) = 211
  Grenze(i) = 911
  nurpath(i) = 2
  i = i + 1
  
  Üs(i) = "Folsre"
  Foli = i
  Titel(i) = "Folsäure [ng/ml], NB [2,6 - 14,6]"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü like 'FOLS%' order by zp desc"
  Fqmin(i) = 0
  UGrenze(i) = 2.6
  Grenze(i) = 14.6
  nurpath(i) = 2
  i = i + 1
  
  Üs(i) = "Leuko"
  Leuki = i
  Titel(i) = "Leukozyten im Blutbild [/nl]"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü in ('LEUK', 'LEUKO') order by zp desc"
  Fqmin(i) = 1
  UGrenze(i) = 4#
  Grenze(i) = 9.4
  nurpath(i) = 2
  i = i + 1
    
  Üs(i) = "CRP"
  CRPi = i
  Titel(i) = "C-reaktives Protein [mg/dl], NB < 5"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü in ('CRP', 'CRPQ') order by zp desc"
  Fqmin(i) = 0
  Grenze(i) = 5#
  nurpath(i) = 2
  i = i + 1
    
  Üs(i) = "CK"
  CRPi = i
  Titel(i) = "C(P)K [U/l], NB < 190"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü = 'CKNAC' order by zp desc"
  Fqmin(i) = 0
  Grenze(i) = 190
  nurpath(i) = 2
  i = i + 1
    
  Üs(i) = "Carotis"
  Caroi = i
  Titel(i) = "Duplexsonographie der Halsschlagadern"
  sql(i) = "select zeitpunkt zp, inhalt wert from eintraege where pat_id = " & pid & " and art like 'duplex%' and (inhalt like '%Halsschlag%' or inhalt like '%Carotis%' or inhalt like '%ACC%' or inhalt like '%ACI%' or inhalt like '%subcl%') order by zp desc"
  Fqmin(i) = 0.5
  i = i + 1
    
  Üs(i) = "Car alle"
  CRPi = i
  Titel(i) = "Duplexsonographie der Halsschlagadern alle Textstellen"
  sql(i) = "select zeitpunkt zp, inhalt wert from eintraege where pat_id = " & pid & " and (inhalt like '%Halsschlag%' or inhalt like '%Carotis%') order by zp desc"
  Fqmin(i) = 0
  nurpath(i) = 2
  i = i + 1
    
    
  For i = 0 To zl
   rs(i).Open sql(i), DBCn, adOpenStatic, adLockReadOnly
  Next i
  
' Grenze für Blutdruck korrigieren
  If Not rs(Albi).BOF Then
   Do While Not rs(Albi).EOF
    Wert = rs(Albi)!Wert
    WertNum = MachNumerisch(Replace(Wert, ".", ","))
    If WertNum > Grenze(Albi) Then pathAlbZ = pathAlbZ + 1
    AlbZ = AlbZ + 1
    If AlbZ = 3 And pathAlbZ = 0 Then Exit Do ' Wenn die letzten 3 in Ordnung sind, dann o.k.
    rs(Albi).Move 1
   Loop
   rs(Albi).MoveFirst
   If pathAlbZ > 1 And pathAlbZ * 3 > AlbZ Then obNP = True
   If Not obNP Then
    If Not rs(Kreai).BOF Then
     Do While Not rs(Kreai).EOF
      Wert = rs(Kreai)!Wert
      WertNum = MachNumerisch(Replace(Wert, ".", ","))
      If WertNum > Grenze(Kreai) Then pathKreaZ = pathKreaZ + 1
      KreaZ = KreaZ + 1
      If KreaZ = 3 And pathKreaZ = 0 Then Exit Do
      rs(Kreai).Move 1
     Loop
     rs(Kreai).MoveFirst
    End If
    If pathKreaZ > 1 And pathKreaZ * 3 > KreaZ Then obNP = True
   End If ' Not obNP Then
  End If
  If obNP Then
   Grenze(RRi) = 120
  End If
  Dim üdt As dmptyp
  Dim t0#, T1#
  t0 = Timer
  Call DMPString(CLng(pid), üdt, True)
  T1 = Timer
  If üdt.fußst = auff Or üdt.sens = auff Or üdt.puls = auff Or üdt.FEn(12) = True Then
   Fqmin(Neuri) = 4
   rsMB.Open "select * from eintraege where pat_id = " & pid & " and art = 'notiz' and inhalt = 'Merkblatt Fußsyndrom mitgegeben'", DBCn, adOpenStatic, adLockReadOnly
   If rsMB.EOF Then obMB = True
  End If
' ob noch Avandia gegeben wird
  Set rsMB = Nothing
  rsMB.Open "select zeitpunkt zp, medikament wert from (" & lMP & ") as innen where medikament like 'Avan%'", DBCn, adOpenStatic, adLockReadOnly
  If Not rsMB.EOF Then obAvan = True
  Set rsMB = Nothing
  
' Wenn keine Diagnose Hypertonie und Blutdruckwerte normal, dann keinen RR-Vergleich
  Set rDia = Nothing
  rDia.Open "select * from diagnosen where pat_id = " & pid & " and icd like 'i10%' and diagsicherheit in ('g',' ')", DBCn, adOpenStatic, adLockReadOnly
  If rDia.EOF Then
   If pathz(RRi) < 1 Or (npathz(RRi) > 2 And pathz(RRi) < 2) Then
    nurpath(RRVgli) = 2 ' => RRVergleich wird nicht als fehlend angezeigt
    Fqmin(RRi) = 4 ' => Blutdruck muss nur vierteljährlich gemessen werden
   End If
  End If
    
' Wenn therapieformbezogen zu wenig geschult, dann schulen
  Dim sollz%, obnach%
  Select Case rAn!TherAkt
   Case "CSII", "ICT", "CT", "Komb": sollz = 6
   Case Else: sollz = 4
  End Select
  
' Nachschulungen festlegen
  If Not rs(Schuli).BOF Then
   If QEnd(ZQuart(Now)) - QAnf(ZQuart(rs(Schuli)!Zp)) > 365 Then
    obnach = True
   End If
  End If
    
  rszmax = 0
  For i = 0 To zl
   Set rszahl = Nothing
   rszahl.Open "select count(0) as ct from (" & sql(i) & ") as innen", DBCn, adOpenStatic, adLockReadOnly
   rsz(i) = rszahl!ct
' Hier Zahl der Schulungen auswerten
   If i = Schuli Then
    If rsz(i) < sollz Or obnach Then
     Fqmin(Schuli) = 365
    End If
   End If
   If rsz(i) > zmax(i) Then rsz(i) = zmax(i)
   If rsz(i) > rszmax Then rszmax = rsz(i)
   
'   If i = Gwi Then
'    If Not rs(i).BOF Then
'     If rAn!Größe <> 0 Then
'      bmi = MachNumerisch(rAn!Größe)
'      If bmi > 3 Then bmi = bmi / 100
'      bmi = Round(MachNumerisch(rs(i)!Wert) / (bmi * bmi), 1)
'      bmiS = bmi & " kg/m²"
'     End If
'    End If
'   End If
   
   If Not rs(i).BOF Then
    loopct = 1
    Do While Not rs(i).EOF
     If UGrenze(i) <> -1 Or Grenze(i) <> -1 Then
' allgemein (für alle Parameter) festlegen, ob pathologisch
      Wert = rs(i)!Wert
      WertNum = MachNumerisch(Replace(Wert, ".", ","))
      If (i = Urini And InStrB(Wert, "++") <> 0 Or InStrB(Wert, "000") <> 0) Or _
      (UGrenze(i) <> -1 And WertNum < UGrenze(i)) Or (Grenze(i) <> -1 And WertNum > Grenze(i)) Then
       pathol(i, loopct) = True
       pathol(i, 0) = True
      End If
     End If

' bei pathologischen Laborwerten auch andere anzeigen
     If pathol(i, 0) Then
      If i = Hbi Then
       nurpath(Ferri) = 3
       nurpath(B12i) = 3
       nurpath(Foli) = 3
      End If
      If i = Leuki Then
       nurpath(CRPi) = 3
      End If
      If i = TSHi Then
       nurpath(fT4i) = 3
       nurpath(fT3i) = 3
      End If
      If i = TGi Then
       nurpath(GGTi) = 0
       nurpath(GPTi) = 0
      End If
      If i = Kreai Then
       nurpath(GFRi) = 3
      End If
      If i = Caroi Then
       nurpath(Caroai) = 3
      End If
     End If ' pathol(i, 0) Then
    
' auch diastolischen Wert berücksichtigen
     If i = RRi Then
      Wert = getDiast(rs(i)!Wert)
      If IsNumeric(Wert) Then
       If CDbl(Wert) >= GrenzeDiast Then
        pathol(i, loopct) = True
        pathol(i, 0) = True
       End If
      End If
     End If
     If pathol(i, loopct) Then
      pathz(i) = pathz(i) + 1
     Else
      npathz(i) = npathz(i) + 1
     End If
     rs(i).Move 1
     loopct = loopct + 1
     If loopct > zmaxges Then Exit Do
    Loop
    rs(i).MoveFirst
   End If ' not rs(i).bof
  Next i
  
'Beginn mit Schreiben der Daten
  Print #355, "<h1>" ' , <span" & IIf(bmi > 24.9, " id='path'", "") & ">" & bmiS & "</span>
  Print #355, " <B>" & rNam!Vorname & " " & rNam!Nachname & "</B>, *" & Format(rNam!gebdat, "d.m.yy") & " (" & CInt((Now - rNam!gebdat) / 365) & "a), <span style='color:blue'><span id='unauff'>&nbsp;&nbsp;Pat_id: </span>" & pid & "</span><span id = 'unauff'>,&nbsp;&nbsp;&nbsp;vorgestellt: " & Format(rAn!Vorgestellt, "d.m.yy") & ",&nbsp;&nbsp;&nbsp;für: </span>" & Format(Datum, "d.m.yy") & " <span style='font-size:smaller'>" & Format(Uhrzeit, "hh:mm") & ",</span>&nbsp;&nbsp;&nbsp;<span id='unauff'>" & rNam!Notiz & ",&nbsp;&nbsp;&nbsp;Therapie zuletzt: </span>" & rAn!TherAkt & "</h1>" ' TherapieArtEinzelnFestlegen(CLng(pid), rAn) & "</span></h1>"
  Print #355, "</h1>"
'  If Not obhierdmp(rNam!Notiz) Then
  Dim obdmp%
  Call obhierdmp(rNam!Notiz, , , , obdmp)
  If Not obdmp Then
   If Not rFl.EOF Then
    If rFl!Kateg <> "LKK" Then
     Print #355, "<div id='unauff'><span id = 'cave'>Nicht im DMP</span>: " & IIf(rNam!Notiz = "", "(&Oslash;&nbsp;Eintrag)", rNam!Notiz) & "</div>"
    End If
   End If
  End If
  Print #355, "</h1>"
  Print #355, ""
  
  Call getHausarzt(CLng(pid), üw12erg)
  obpath = -1
  For i = 0 To UBound(üw12erg, 2)
   If InStrB(üw12erg(10, i), "HA") <> 0 Then
    obpath = 0
    Exit For
   End If
  Next i
  Print #355, "<div id='unauff'" & IIf(obpath, " id='cave'", "") & ">"
  For i = 0 To UBound(üw12erg, 2)
   If üw12erg(10, i) <> "" Then
    Print #355, "<span title='" & üw12erg(1, i) & ", " & üw12erg(2, i) & ", " & üw12erg(3, i) & ", " & üw12erg(8, i) & ", Tel: " & üw12erg(13, i) & ", Fax: " & üw12erg(4, i) & "'>"
    Print #355, üw12erg(10, i) & ": " & üw12erg(1, i) & " (T:" & üw12erg(13, i) & "), "
    Print #355, "</span>"
   End If
  Next i
  Print #355, "</div>"
  If obpath Then Print #355, "<div id='cave'>Hausarzt nicht in Turbomed eingetragen</div>"
  
  Set rFlb = Nothing
  rFlb.Open "SELECT dokumente.pat_id, dokumente.zeitpunkt as zp, dokumente.dokpfad, dokumente.dokname, dokumente.quelldatum, dokumente.dokgroe, kvnr, übwv, if(`dokumente abgehakt`.abgehakt,'X',' ') as abgeh FROM ((dokumente LEFT JOIN `dokumente abgehakt` ON dokumente.DokPfad = `dokumente abgehakt`.DokPfad) left join namen on dokumente.pat_id = namen.pat_id) left join `letzte faelle` on dokumente.pat_id = `letzte faelle`.pat_id WHERE (((dokumente.DokName) Like '%Fremdlabor%')) and (isnull(abgehakt) or abgehakt = 0) and dokumente.pat_id = " & pid & " order by dokumente.pat_id desc, zp desc", DBCn, adOpenDynamic, adLockReadOnly
  If Not rFlb.BOF Then
   Print #355, "<table>"
   Print #355, "  <tr>"
   Print #355, "<p id='abstand'> </p>"
   Print #355, "   <th id='cave'>Fehlende Fremdlabore:</th>"
   FlbZ = 0
   Do
    Print #355, "    <td><p>" & Format(rFlb!Zp, "d.m.yy") & "</p></td>"
    rFlb.MoveNext
    If rFlb.EOF Then Exit Do
    FlbZ = FlbZ + 1
    If FlbZ > 4 Then Exit Do
   Loop
   Print #355, "  </tr>"
   Print #355, "</table>"
  End If ' not rFlb.bof
    
' Liste der ganz fehlenden Parameter
  For i = 0 To zl
   Dim obschonabstand%
   If rsz(i) = 0 Then
    If Not obschonabstand Then
     Print #355, "<p id='abstand'> </p>"
     obschonabstand = True
    End If
    If nurpath(i) < 2 Then
     Print #355, "     <span id='cave'>&Oslash;&nbsp;"; Üs(i) & ",&nbsp;&nbsp;</span>"
    End If
   End If
  Next i
  If obMB Then Print #355, "    <span id='cave'>&Oslash;&nbsp;Merkblatt Fußsyndrom,&nbsp;&nbsp;</span>"
  If obAvan Then Print #355, "    <span id='cave'>&nbsp;Avandia,&nbsp;&nbsp;</span>"
  
  ie = -1
  For k = 0 To 5 ' Blöcke
   ia = ie + 1
   ie = ie + 9
   If ie > zl Then ie = zl
   Print #355, "<p id='abstand'> </p>"
   Print #355, "<TABLE>"
'Überschriften
   Print #355, "   <TR>"
   For i = ia To ie
    If nurpath(i) = 0 Or nurpath(i) = 3 Or pathol(i, 0) Then
     If rsz(i) <> 0 Then
      Print #355, "     <th title='" & Titel(i) & "' colspan='2'" & IIf(pathol(i, 0) Or (i = Gwi And üdt.bmi > 24.9), " id='path'", " id='schmal'") & ">" & Üs(i) & IIf(i = RRi, " (Ziel: &le;" & Grenze(i) & "/" & GrenzeDiast & ")", "") & IIf(i = Neuri, " (" & IIf(üdt.fußst = auff, "Fußst,", "") & IIf(üdt.sens = auff, "Sens,", "") & IIf(üdt.puls = auff, "Puls,", "") & IIf(üdt.FEn(12), "<span id='cave'>DFS</span>", "") & "&rArr;" & Fqmin(i) & "/a)", "") & IIf(i = Gwi, " (BMI " & Round(üdt.bmi, 1) & " kg/m²)", "") & IIf(i = Schuli, " (Soll:" & sollz & IIf(obnach, "+", "") & ")", "") & "</th>"
     End If
    End If
   Next i
   Print #355, "   </TR>"
   For j = 1 To rszmax
'Werte(j)
    Print #355, "   <TR>"
    For i = ia To ie ' Parameter in jedem Block
     If i <> SchulEintri Or Fqmin(Schuli) = 365 Then
      If nurpath(i) = 0 Or nurpath(i) = 3 Or pathol(i, 0) Then
       If rsz(i) >= j Then
        obpath = 0
        If j = 1 Then
         Dim Zp As Date
         Zp = QAnf(ZQuart(rs(i)!Zp))
' wenn z.B. nicht in diesem Quartal (fqmin(i)=4) und auch nicht jünger als die Hälfte des entsprechenden Abstandes, dann rot
         If Fqmin(i) <> 0 Then
          If Now - Zp > 365 / Fqmin(i) And (Now - rs(i)!Zp) > 365 / Fqmin(i) * 0.5 Then obpath = True
         End If ' Fqmin(i) <> 0 Then
        End If ' j = 1 Then
        Print #355, "    <td id='schmal'><P" & IIf(obpath, " id='cave'", "") & ">" & Format(rs(i)!Zp, "d.m.yy") & "</td>"
        Print #355, "    <td" & IIf(pathol(i, j), " id='path'", "") & "><p>" & rs(i)!Wert & "</p></td>"
        rs(i).MoveNext
       Else  ' rsz(i) >= j Then
        If rsz(i) <> 0 Then
         Print #355, "    <td id='schmal'></td><td></td>"
        End If ' rsz(i) <> 0 Then
       End If ' rsz(i) >= j Then
      End If ' nurpath(i) = 0 Or pathol(i, 0) Then
     End If ' i <> SchulEintri Or Fqmin(Schuli) = 365 Then
    Next i
    Print #355, "   </TR>"
    If j = rszmax Then
     If i = zl Then
      Exit For ' fertig
     End If
    End If
   Next j
   Print #355, "</TABLE>"
  Next k
  Print #355, "<p id='abstand'> </p>"
  Print #355, "<div id='abstand'>" & Replace(Replace(DiagString(CLng(pid)), vbVerticalTab & " ", "<br />&nbsp;"), vbVerticalTab, "<br />") & "</div>"
  Print #355, "</BODY>"
  Print #355, "</HTML>"
  
  Close #355
  DoEvents
  Call SetFileTimeByDate(DateiName, mftLastWriteTime, Datum + Uhrzeit)
  Lese.Ausgabe = "Patientenlaufzettel für " & rNam!Nachname & ", " & rNam!Vorname & " (Pat_id: " & pid & ") erstellt." & vbCrLf & altAusgabe
 End If ' not rnam.eof
 altAusgabe = Lese.Ausgabe
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dodoPLZ/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dodoPLZ

Function doPLZeinzeln()
 Dim erg$
 erg = InputBox("Bitte Pat-ID aus Turbomed eingeben:", "Patientenlaufzettel Eingabe")
 If IsNumeric(erg) Then
  Call dodoPLZ(erg, Now, Now - Int(Now))
 End If
End Function

Public Function bittest1()
Set DBCn = Nothing
If DBCn Is Nothing Or DBCn.ConnectionString = "" Then Call acon(quelleT)
Dim rs As New ADODB.Recordset
rs.Open "select * from test.test", DBCn, adOpenStatic, adLockReadOnly
Debug.Print rs!Bit, DBCn
rs.Move 1
Debug.Print rs!Bit, DBCn
End Function
