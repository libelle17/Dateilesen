Attribute VB_Name = "Laufzettel"
Option Explicit
Public plzVerz$
'Const plzVerz$ = pVerz & "plz\"

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
 Dim Str As New CString, NTL&
 Dim i%, j&, splitt$(), s2$(), Pid$, Datum As Date
 Dim Uhrzeit$
 Filename = plzVerz & "TMFTools.pdf"
' Filename = pverz & "Hausmann Karteikarte.pdf"
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
   Str.Append PDTextS.GetText(i)
  Next i
  DoEvents
  Str.Append vbCrLf
  j = j + 1
 Loop
 Result = PDDoc.Close
 Set PDPage = Nothing
 Set PDHili = Nothing
 Set PDTextS = Nothing
 Set PDDoc = Nothing
 SplitNeu Str.Value, vbCrLf, splitt
 SplitNeu splitt(0), " ", s2
 For j = 0 To UBound(s2)
  If InStrB(s2(j), "-") > 0 Then
   s2 = Split(s2(j), "-")
   Datum = CDate(s2(0))
   Exit For
  End If
 Next j
 erg = Dir$(plzVerz & "*")
 Do While erg <> ""
  If InStrB(erg, "Patientenlaufzettel") <> 0 Then
   Kill plzVerz & erg
  End If
  erg = Dir
 Loop
 DoEvents
 For j = 0 To UBound(splitt)
  If InStrB(splitt(j), " ") <> 0 Then
   i = InStr(splitt(j), " ")
   Pid = Left$(splitt(j), i - 1)
   If IsNumeric(Pid) Then
    If InStrB(Pid, ".") = 0 And InStrB(Pid, ",") = 0 And InStrB(Pid, ";") = 0 Then
     Uhrzeit = ""
     i = InStr(splitt(j), ":")
     If i <> 0 Then
      Uhrzeit = Mid(splitt(j), i - 2, 5)
     End If
     Call dodoPLZ(CLng(Pid), Datum, Uhrzeit)
    End If
   End If
  End If
 Next j
 MsgBox "Fertig!"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doPatientenlaufzettel/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
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
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in getDiast/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' getDiast
Function dodoPLZ(Pid$, Optional Datum As Date, Optional Uhrzeit$, Optional mitanzeig%) ' Patientenlaufzettel
 Dim rnam As New ADODB.Recordset, rAn As New ADODB.Recordset, rFl As New ADODB.Recordset, rFlb As New ADODB.Recordset, FlbZ&
 Dim rDia As New ADODB.Recordset, rsMB As New ADODB.Recordset, rsDi As New ADODB.Recordset
 Dim rHAV As New ADODB.Recordset
 Dim lsql$, lMP$, i&, Größe!, bmi!, bmiS$, Gewicht!, Wert$, WertNum!, j&, k&, ia&, ie&
 Const zl% = 47     ' Zahl der Parameter
 Const SchulPID& = 1789
 Const SpZahl% = 11
 Dim RRi% ' Parameter für Blutdruck
 Dim Neuri% ' Parameter für Neurostatus
 Dim Fußi% ' Parameter für Fußinspektion
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
 Dim DateiName$, DateiNameAkt$
 Dim obMB% ' ob Merkblatt Fußsyndrom mitgegeben
 Dim obAvan% ' ob Avandia verordnet wird
 Dim rs(zl) As New ADODB.Recordset, rszahl As New ADODB.Recordset
 Dim sql$(zl)
 Dim rsz&(zl) ' vorhandene Zahl
 Dim zmax%(zl), rszmax& ' maximal dargestellte Zahl an Ausprägungen jedes Parameters, Maximum davon je Großzeile
 Dim Weite$(zl) ' Angabe zur Weite
 Dim RowSp%(zl) ' Zahl der Zeilen pro Kleinzeile
 Dim Pause%(zl) ' Zahl der noch zu pausierenden Zeilen für die aktuelle Spalte wegen Rowspan vorher
 Dim Üs$(zl), Titel$(zl)  ' Überschrift
 Dim Fqmin(zl) ' Mindest-Untersuchungs-Zahl pro Jahr
 Dim UGrenze!(zl), Grenze!(zl) ' Grenze für pathologische Werte
 Dim obpath%(zl)
 Dim KVNr&
 Const GrenzeDiast! = 81
 Dim nurpath%(zl) ' -1 = nicht anzeigen
                  ' 1 = Anzeige nur bei Fehlen eines Wertes und bei pathologischem Wert,
                  ' 2 = Anzeige nur bei pathologischem Wert
                  ' 3 = Anzeige bei Vorhandensein eines Wertes
 Dim loopct%, pathol%(zl, zmaxges) ' pathol(zl,0)= ' ob irgend einer der Werte pathologisch ist
 Dim pathz%(zl), npathz%(zl) ' Zahl der nicht/ Pathologischen
 Dim üw12erg$()
 
 On Error GoTo fehler
 If LenB(DBCn) = 0 Or LenB(DBCn.ConnectionString) = 0 Then Call acon(quelleT)
 rnam.Open "select * from namen where pat_id = " & Pid, DBCn, adOpenStatic, adLockReadOnly
' If PID = 2183 Then Stop
 rFl.Open "select k.*, f.* from faelle f join aktfv a using (fid) left join kassenliste k on(((f.VKNr = k.VK) and (f.IK = k.IK))) where f.pat_id = " & Pid & " group by f.pat_id", DBCn, adOpenStatic, adLockReadOnly
 If Not rnam.EOF Then
  rAn.Open "select * from anamnesebogen where pat_id = " & Pid, DBCn, adOpenStatic, adLockReadOnly
  If Datum = 0 Then Datum = Now
  DateiName = plzVerz & rnam!Nachname & " " & rnam!Vorname & ",   Pid " & rnam!Pat_id & ", Patientenlaufzettel"  ' replace$(CDate(Uhrzeit), ":", ".") & " " & Format(Datum, "dd.mm.yyyy") &
'  Dim i%
  On Error Resume Next
  For i = 0 To 10
   Err.Clear
   DateiNameAkt = DateiName & String(i, "_") & ".html"
   Open DateiNameAkt For Output As #355
   If Err.Number = 0 Then Exit For
  Next i
  On Error GoTo fehler
  Print #355, "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>"
  Print #355, "<HTML>"
  Print #355, "<HEAD>"
  Print #355, "    <META HTTP-EQUIV='CONTENT-TYPE' CONTENT='text/html; charset=windows-1252'>"
  Print #355, "    <TITLE>" & rnam!Nachname & " " & rnam!Vorname & " (Pat_id:" & Pid & "): Patientenlaufzettel vom " & Now & "</TITLE>"
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
  Print #355, "      #schmal { width:2.5em;border-left-width:1.5mm;padding:0cm; }"
  Print #355, "      #lnn { width:2.5em;border-left-width:1.5mm;padding:0cm; }"
  Print #355, "      #lno { width:2.5em;border-left-width:1.5mm;padding:0cm;color:blue; }"
  Print #355, "      #lpn { width:2.5em;border-left-width:1.5mm;padding:0cm;font-weight:bold;background-color:yellow;font-size:2.8mm; }"
  Print #355, "      #lpo { width:2.5em;border-left-width:1.5mm;padding:0cm;font-weight:bold;color:blue;background-color:yellow;font-size:2.8mm; }"
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
  lsql = "select * from labor1 where pat_id = " & Pid & " union select * from labor2 where pat_id = " & Pid
  lMP = "select mp.* from (select pat_id, mpnr from (select mpi.pat_id, mpi.mpnr, mpi.zeitpunkt from (select pat_id, max(zeitpunkt) as zp from medplan mp where pat_id = " & Pid & " and zeitpunkt <= now() group by pat_id) as mpl left join medplan mpi on mpl.pat_id = mpi.pat_id and mpl.zp = mpi.zeitpunkt group by mpl.pat_id, mpl.zp, mpnr order by  mpl.pat_id, mpl.zp, mpnr desc) as mpii group by pat_id, zeitpunkt) as innen left join medplan mp on innen.pat_id = mp.pat_id and innen.mpnr = mp.mpnr"
  i = 0
  
  Üs(i) = "Gw"
  Titel(i) = "Körpergewicht [kg]"
  Gwi = i
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where art like 'gewicht%' and pat_id = " & Pid & " order by zp desc"
  Fqmin(i) = 4
  i = i + 1
  
  Üs(i) = "Taille"
  Titel(i) = "Taille (Mitte zwischen Hüfte und Beckenkamm) [cm], NB &#x2640;(-80)-88, &#x2642;(-94)-102"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where art like 'taille%' and pat_id = " & Pid & " order by zp desc"
  Fqmin(i) = 1
  If rnam!Geschlecht = "m" Then
   Grenze(i) = 102
  Else
   Grenze(i) = 88
  End If
  i = i + 1
  
  Üs(i) = "RR"
  Titel(i) = "Blutdruck [mm Hg]<br>Ziel bei Nephropathie niedriger"
  RRi = i
  sql(i) = "select zeitpunkt zp, rr wert from rr where pat_id = " & Pid & " order by zeitpunkt desc"
  Fqmin(i) = 180
  Grenze(i) = 130
  i = i + 1
  
  Üs(i) = "HbA1c"
  Titel(i) = "HbA1c [%], nach DCCT, NB [4,1-6,2]"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü like 'HBA1C%' order by zp desc"
  Fqmin(i) = 4
  Grenze(i) = 6.5
  i = i + 1
  
  Üs(i) = "Uzu"
  Titel(i) = "Unterzucker seit letztem Besuch"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where (art like 'uzu%' or art like 'hypo%') and pat_id = " & Pid & " order by zp desc"
  Fqmin(i) = 4
  i = i + 1
  
  Üs(i) = "Hypo"
  Titel(i) = "Schwere Hypoglykämie"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where (art like 'hypo%') and pat_id = " & Pid & " order by zp desc"
  nurpath(i) = 3
  Fqmin(i) = 0
  i = i + 1
  
  Üs(i) = "Hyper"
  Titel(i) = "Hyerglykämische Entgleisung"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where (art like 'hyper%') and pat_id = " & Pid & " order by zp desc"
  If Left$(rAn!Diabetestyp, 1) = "1*" Then
   nurpath(i) = 3 '1
  Else
   nurpath(i) = 3
  End If
  Fqmin(i) = 0
  i = i + 1
  
  Üs(i) = "NeuSt"
  Titel(i) = "Neuro-Status"
  Neuri = i
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where art like 'usdm%' and pat_id = " & Pid & " order by zp desc"
  Fqmin(i) = 1
  zmax(i) = 1
  Weite(i) = "20%"
  RowSp(i) = 8
  i = i + 1
  
  Üs(i) = "Fuß"
  Titel(i) = "Fußuntersuchung"
  Fußi = i
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where art like 'fuß%' and pat_id = " & Pid & " order by zp desc"
  Fqmin(i) = 0
  zmax(i) = 4
  RowSp(i) = 2
  i = i + 1
  
  Üs(i) = "Beweg"
  Titel(i) = "Bewegungsanamnese"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where (art like 'beweg%') and pat_id = " & Pid & " order by zp desc"
  Fqmin(i) = 1
  i = i + 1
  
  Üs(i) = "Beruf"
  Titel(i) = "Berufsanamnese"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where (art like 'beruf%') and pat_id = " & Pid & " order by zp desc"
  Fqmin(i) = 0
  i = i + 1
  
  Üs(i) = "Auto"
  Titel(i) = "Aufklärung über Autofahrverbot"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where (art like 'auto%') and pat_id = " & Pid & " order by zp desc"
  Fqmin(i) = 0
  i = i + 1
  
  Üs(i) = "Keto"
  Titel(i) = "Aufklärung über Ketonkörpermessung"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where (art like 'keto%') and pat_id = " & Pid & " order by zp desc"
  Fqmin(i) = 0
  i = i + 1
  
  Üs(i) = "K+"
  Titel(i) = "Kalium [mmol/l], NB 3.5 - 5.6"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where (abkü = 'K' or abkü like 'KALI%') order by zp desc"
  UGrenze(i) = 3.5
  Grenze(i) = 5.6
  Fqmin(i) = 1
  nurpath(i) = 2
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
  nurpath(i) = 1
  i = i + 1
  
  Üs(i) = "GGT"
  Titel(i) = "Gamma-GT [U/l], NB &#x2640;-39, &#x2642;-66"
  GGTi = i
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü like '%GT%' order by zp desc"
  Fqmin(i) = 1
  If rnam!Geschlecht = "m" Then
   Grenze(i) = 66
  Else
   Grenze(i) = 39
  End If
  nurpath(i) = 1
  i = i + 1
  
  Üs(i) = "GPT"
  Titel(i) = "GPT (ALAT) [U/l], NB &#x2640;-35, &#x2642;-50"
  GPTi = i
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü like 'GPT%' order by zp desc"
  Fqmin(i) = 1
  If rnam!Geschlecht = "m" Then
   Grenze(i) = 50
  Else
   Grenze(i) = 35
  End If
  nurpath(i) = 1
  i = i + 1
  
  Üs(i) = "Ferr"
  Titel(i) = "Ferritin [ng/ml], NB &#x2640;15-150, &#x2642;30-400"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü like 'FERR%' order by zp desc"
  If rnam!Geschlecht = "m" Then
   UGrenze(i) = 30
   Grenze(i) = 400
  Else
   UGrenze(i) = 15
   Grenze(i) = 150
  End If
  Fqmin(i) = 0
  nurpath(i) = 2
  i = i + 1
  
  Üs(i) = "Digo"
  Titel(i) = "Digoxin [ng/ml], NB 0,8-2,0"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü like 'DIGO%' order by zp desc"
  UGrenze(i) = 0.8
  Grenze(i) = 2#
  Fqmin(i) = 0
  nurpath(i) = 2
  i = i + 1
  
  Üs(i) = "Digit"
  Titel(i) = "Digitoxin [ng/ml], NB 10-30"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü like 'DIGI%' order by zp desc"
  UGrenze(i) = 10#
  Grenze(i) = 30#
  Fqmin(i) = 0
  nurpath(i) = 2
  i = i + 1
  
  Üs(i) = "Urin"
  Titel(i) = "Urinbefund im Teststreifen"
  Urini = i
  sql(i) = "select * from (select zeitpunkt zp, inhalt wert from eintraege where pat_id = " & Pid & " and art like 'urin%') u1 union select * from (select eingang zp, keimzahl wert from laborxbakt natural join laborxus where pat_id = " & Pid & " and keimzahl<> '') u2 order by zp desc"
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
  If rnam!Geschlecht = "m" Then
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
  sql(i) = "select * from (select zeitpunkt zp, inhalt wert from eintraege where pat_id = " & Pid & " and (" + artinartspez + " and ((inhalt like '%augenb%' and not inhalt like '%augenbl%' and not inhalt like '%augen') or (inhalt like '%augenarzt%' or inhalt like '%augenärzt%') or (inhalt like '% aa%' and not inhalt like '% aag%'))) or (art = 'aa' or art = 'augen'))" & " and zeitpunkt > " & DatForm(Now() - 550) & ") u1 " & _
     "union select * from (select zeitpunkt zp, dokname wert from dokumente where pat_id = " & Pid & " and dokname like '%augen%' and zeitpunkt > " & DatForm(Now() - 550) & ") u2 " & _
     "union select * from (select vorgestellt zp, concat(`Augensp zuletzt`,': ',`Augensp Befund`) wert from anamnesebogen where pat_id = " & Pid & " and not isnull(`Augensp zuletzt`)) u3 " & _
     "order by zp desc"
  Fqmin(i) = 1
  i = i + 1
  
  Üs(i) = "RR-Vgl"
  Titel(i) = "Vergleich des Blutdruckgerätes mit unserem Oberarmmanschettengerät"
  RRVgli = i
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where art like 'rrvgl%' and pat_id = " & Pid & " order by zp desc"
  Fqmin(i) = 1
  zmax(i) = 3
  Weite(i) = "7em"
  RowSp(i) = 2
  i = i + 1
  
  Üs(i) = "BZ-Vgl"
  Titel(i) = "Vergleich des Blutzuckergerätes mit Naßchemiegerät"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where art like 'bzvgl%' and pat_id = " & Pid & " order by zp desc"
  Fqmin(i) = 1
  zmax(i) = 3
  Weite(i) = "7em"
  RowSp(i) = 2
  i = i + 1
  
  Üs(i) = "Schul"
  Schuli = i
  Titel(i) = "(Gruppen-)Schulungen, Nachschulungen"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where art like 'schul%' and pat_id = " & Pid & " order by zp desc"
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
  sql(i) = "select zeitpunkt zp, concat (art, ' <- ', inhalt) wert from eintraege where pat_id = 1789 and inhalt like '%" & rnam!Nachname & "%' order by zp desc"
  nurpath(i) = 3
  i = i + 1
  
  Üs(i) = "Hb"
  Titel(i) = "Hämoglobin [g/dl], NB &#x2640;12-16, &#x2642;14-18"
  Hbi = i
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü in ('Hb') order by zp desc"
  Fqmin(i) = 2
  If rnam!Geschlecht = "m" Then
   UGrenze(i) = 14
   Grenze(i) = 18
  Else
   UGrenze(i) = 12
   Grenze(i) = 16
  End If
  nurpath(i) = 1
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
'  cki = i
  Titel(i) = "C(P)K [U/l], NB < 190"
  sql(i) = "select zeitpunkt zp, if(isnull(wert),if(isnull(kommentar),'',kommentar),wert) wert from (" & lsql & ") as innen where abkü = 'CKNAC' order by zp desc"
  Fqmin(i) = 0
  Grenze(i) = 190
  nurpath(i) = 2
  i = i + 1
    
  Üs(i) = "A.P."
  Titel(i) = "Bericht über Angina pectoris"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where (art = 'ap') and pat_id = " & Pid & " order by zp desc"
  Fqmin(i) = 0
  i = i + 1
  
  Üs(i) = "Carotis"
  Caroi = i
  Titel(i) = "Duplexsonographie der Halsschlagadern"
  sql(i) = "select zeitpunkt zp, inhalt wert from eintraege where pat_id = " & Pid & " and (art like 'caro%' or art like 'duplex%' and (inhalt like '%Halsschlag%' or inhalt like '%Carotis%' or inhalt like '%ACC%' or inhalt like '%ACI%' or inhalt like '%subcl%')) order by zp desc"
  Fqmin(i) = 0.5
  RowSp(i) = 2
  zmax(i) = 4
  i = i + 1
    
  Üs(i) = "Car alle"
  Caroai = i
  Titel(i) = "Duplexsonographie der Halsschlagadern alle Textstellen"
  sql(i) = "select zeitpunkt zp, inhalt wert from eintraege where pat_id = " & Pid & " and (inhalt like '%Halsschlag%' or inhalt like '%Carotis%') order by zp desc"
  Fqmin(i) = 0
  nurpath(i) = 2
  i = i + 1
    
  Üs(i) = "Impf"
  Titel(i) = "Impfaufklärung (Grippe, Pneumonie)"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where (art like 'impf%') and pat_id = " & Pid & " order by zp desc"
  Select Case Month(Now())
   Case 9, 10, 11, 12
    Fqmin(i) = 1
   Case Else
    Fqmin(i) = 0
  End Select
  i = i + 1
    
  Üs(i) = "Colo"
  Titel(i) = "Aufklärung über Darmkrebsfrüherkennung"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where (art like 'colo%') and pat_id = " & Pid & " order by zp desc"
  Dim Alter#
  Alter = (Now() - rnam!GebDat) / 365.24
  If Alter > 54 And Alter < 78 Then
   Fqmin(i) = 1
  Else
   Fqmin(i) = 0
  End If
  i = i + 1
    
  Üs(i) = "Pros"
  Titel(i) = "Aufklärung über Urologische Früherkennung"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where (art like 'pros%') and pat_id = " & Pid & " order by zp desc"
  If rnam!Geschlecht = "m" And Alter > 44 And Alter < 78 Then
   Fqmin(i) = 1
  Else
   Fqmin(i) = 0
  End If
  i = i + 1
    
  Üs(i) = "Gyn"
  Titel(i) = "Aufklärung über gynäkologische Untersuchung"
  sql(i) = "select zeitpunkt zp, inhalt wert FROM eintraege where (art like 'gyn%') and pat_id = " & Pid & " order by zp desc"
  If rnam!Geschlecht = "w" And Alter > 19 Then
   Fqmin(i) = 1
  Else
   Fqmin(i) = 0
  End If
  i = i + 1
    
  For i = 0 To zl
   rs(i).Open sql(i), DBCn, adOpenStatic, adLockReadOnly
  Next i
  
' Grenze für Blutdruck korrigieren
  If Not rs(Albi).BOF Then
   Do While Not rs(Albi).EOF
    Wert = rs(Albi)!Wert
    WertNum = MachNumerisch(Replace$(Wert, ".", ","))
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
      WertNum = MachNumerisch(Replace$(Wert, ".", ","))
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
' Merkblatt Fußsyndrom prüfen
  Dim üdt As dmptyp
  Dim t0#, T1#, MerkblattText$
  t0 = Timer
  Call DMPString(CLng(Pid), üdt, True)
  T1 = Timer
  If üdt.fußst = auff Or üdt.sens = auff Or üdt.sens = pathdok Or üdt.puls = auff Or üdt.FEn(12) = True Then
   Fqmin(Fußi) = 4
' Wenn in diesem Quartal Neurostatus gemacht oder fällig, dann kein Fußstatus nötig
   If Not rs(Neuri).EOF Then
    If ZQuart(rs(Neuri)!Zp) = ZQuart(Now()) Then
     Fqmin(Fußi) = 0
    End If
   End If
   rsMB.Open "select * from eintraege where pat_id = " & Pid & " and art = 'notiz' and inhalt like 'Merkblatt Fußsyndrom%'", DBCn, adOpenStatic, adLockReadOnly
   If rsMB.EOF Then
    obMB = True
    MerkblattText = "Mbl.Fußsy("
    If üdt.fußst = auff Then MerkblattText = MerkblattText & "Fußst.auff,"
    If üdt.sens = auff Then MerkblattText = MerkblattText & "Sens.auff,"
    If üdt.FEn(12) = auff Or üdt.sens = pathdok Then
     MerkblattText = MerkblattText & "ICD("
     Set rsDi = Nothing
     rsDi.Open "select * from diagnosen where pat_id = " & Pid & " and (icd like 'L89%' or icd like 'M14.6%' or icd = 'G63.2') and not diagsicherheit in ('A','Z')", DBCn, adOpenStatic, adLockReadOnly
     Do While Not rsDi.EOF
      MerkblattText = MerkblattText & rsDi!ICD & ","
      rsDi.Move 1
     Loop
     MerkblattText = Left$(MerkblattText, Len(MerkblattText) - 1) & ")"
    End If
    If üdt.puls = auff Then MerkblattText = MerkblattText & "Puls auff,"
    MerkblattText = Left$(MerkblattText, Len(MerkblattText) - 1) & ")"
   End If
  End If
' ob noch Avandia gegeben wird
  Set rsMB = Nothing
  rsMB.Open "select zeitpunkt zp, medikament wert from (" & lMP & ") as innen where medikament like 'Avan%'", DBCn, adOpenStatic, adLockReadOnly
  If Not rsMB.EOF Then obAvan = True
  Set rsMB = Nothing
  
' Wenn keine Diagnose Hypertonie und Blutdruckwerte normal, dann keinen RR-Vergleich
  Set rDia = Nothing
  rDia.Open "select * from diagnosen where pat_id = " & Pid & " and icd like 'i10%' and diagsicherheit in ('g',' ')", DBCn, adOpenStatic, adLockReadOnly
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
    
' In jeder Großzeile muß zunächst die Zahl der Unterzeilen ermittelt werden, diese wiederum hängt von der
' Menge der in der Großzeile dargestellten Spalten ab
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
      WertNum = MachNumerisch(Replace$(Wert, ".", ","))
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
'      If i = TGi Then
'       nurpath(GGTi) = 0
'       nurpath(GPTi) = 0
'      End If
      If i = Kreai Then
       nurpath(GFRi) = 3
      End If
      If i = Caroi Then
       nurpath(Caroai) = 3
       nurpath(Caroi) = -1
      Else
       nurpath(Caroai) = -1
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
  Print #355, " <B>" & rnam!Vorname & " " & rnam!Nachname & "</B>, *" & Format(rnam!GebDat, "d.m.yy") & " (" & CInt((Now - rnam!GebDat) / 365) & "a), <span style='color:blue'><span id='unauff'>&nbsp;&nbsp;Pat_id: </span>" & Pid & "</span><span id = 'unauff'>,&nbsp;&nbsp;&nbsp;vorgestellt: " & Format(rAn!Vorgestellt, "d.m.yy") & ",&nbsp;&nbsp;&nbsp;für: </span>" & Format(Datum, "d.m.yy") & " <span style='font-size:smaller'>" & Format(Uhrzeit, "hh:mm") & ",</span>&nbsp;&nbsp;&nbsp;<span id='unauff'>" & rnam!Notiz & ",&nbsp;&nbsp;&nbsp;Therapie zuletzt: </span>" & rAn!TherAkt & "</h1>" ' TherapieArtEinzelnFestlegen(CLng(pid), rAn) & "</span></h1>"
  Print #355, "</h1>"
'  If Not obhierdmp(rNam!Notiz) Then
  Dim obdmp%
  Call obhierdmp(rnam!Notiz, , , , obdmp)
  If Not obdmp Then
   If Not rFl.EOF Then
    If rFl!Kateg <> "LKK" Then
     Print #355, "<div id='unauff'><span id = 'cave'>Nicht im DMP</span>: " & IIf(rnam!Notiz = "", "(&Oslash;&nbsp;Eintrag)", rnam!Notiz) & "</div>"
    End If
   End If
  End If
  Print #355, "</h1>"
  Print #355, ""
  
  Call getHausarzt(CLng(Pid), üw12erg)
  obpath(0) = -1
  For i = 0 To UBound(üw12erg, 2)
   If InStrB(üw12erg(10, i), "HA") <> 0 Then
    obpath(0) = 0
    Dim KVNrS$
    KVNrS = üw12erg(12, i)
    If IsNumeric(KVNrS) Then
     KVNr = KVNrS
    End If
    Exit For
   End If
  Next i
  Print #355, "<div id='unauff'" & IIf(obpath(0), " id='cave'", "") & ">"
  For i = 0 To UBound(üw12erg, 2)
   If üw12erg(10, i) <> "" Then
    Print #355, "<span title='" & üw12erg(1, i) & ", " & üw12erg(2, i) & ", " & üw12erg(3, i) & ", " & üw12erg(8, i) & ", Tel: " & üw12erg(13, i) & ", Fax: " & üw12erg(4, i) & "'>"
    Print #355, üw12erg(10, i) & ": " & üw12erg(1, i) & IIf(i < 2, " (T:" & üw12erg(13, i) & IIf(üw12erg(6, i) = "X", ", DMP 2", "") & IIf(üw12erg(7, i) = "X", ", DMP 1", "") & ")", "") & ", "
    Print #355, "</span>"
   End If
  Next i
  Print #355, "</div>"
  If obpath(0) Then Print #355, "<div id='cave'>Hausarzt nicht in Turbomed eingetragen</div>"
  If Not rFl.BOF Then ' 5.3.09
   If rFl!Kateg = "AOK" And (KVNr = "6419153" Or obpath(0) = -1) Then
    Set rHAV = Nothing
    rHAV.Open "select * from eintraege where pat_id = " & Pid & " and art = 'hav'", DBCn, adOpenStatic, adLockReadOnly
    If rHAV.BOF Then
     Print #355, "<div id='cave'>Hiesiger AOK-Patient ohne Hausarztvertrag?</div>"
    End If
   End If
  End If
  
  Set rFlb = Nothing
  rFlb.Open "SELECT dokumente.pat_id, dokumente.zeitpunkt as zp, dokumente.dokpfad, dokumente.dokname, dokumente.quelldatum, dokumente.dokgroe, kvnr, übwv, if(`dokumente abgehakt`.abgehakt,'X',' ') as abgeh FROM ((dokumente LEFT JOIN `dokumente abgehakt` ON dokumente.DokPfad = `dokumente abgehakt`.DokPfad) left join namen on dokumente.pat_id = namen.pat_id) left join `lfaelle` on dokumente.pat_id = `lfaelle`.pat_id WHERE (((dokumente.DokName) Like '%Fremdlabor%')) and (isnull(abgehakt) or abgehakt = 0) and dokumente.pat_id = " & Pid & " order by dokumente.pat_id desc, zp desc", DBCn, adOpenDynamic, adLockReadOnly
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
    If nurpath(i) < 2 And Fqmin(i) <> 0 Then ' letzteres 15.11.08
     Print #355, "     <span id='cave'>&Oslash;&nbsp;"; Üs(i) & ",&nbsp;&nbsp;</span>"
    End If
   End If
  Next i
  If obMB Then Print #355, "    <span id='cave'>&Oslash;&nbsp;" & MerkblattText & ",&nbsp;&nbsp;</span>"
  If obAvan Then Print #355, "    <span id='cave'>&nbsp;Avandia,&nbsp;&nbsp;</span>"
  
  ie = -1
  k = 0
  Do ' Blöcke
   ia = ie + 1
   ie = ie + SpZahl
   If ie > zl Then ie = zl
   Print #355, "<p id='abstand'> </p>"
   Print #355, "<TABLE>"
'Überschriften
   Print #355, "   <TR>"
   i = ia
   Do
    If i > ie Then Exit Do
    If ((i <> SchulEintri Or Fqmin(Schuli) = 365) And ((nurpath(i) = 0 Or nurpath(i) = 3 Or pathol(i, 0)) And rsz(i) <> 0)) Then
      Print #355, "     <th title='" & Titel(i) & "' colspan='2'" & IIf(pathol(i, 0) Or (i = Gwi And üdt.bmi > 24.9), " id='path'", " id='schmal'") & ">" & Üs(i) & IIf(i = RRi, " (Ziel: &le;" & Grenze(i) & "/" & GrenzeDiast & ")", "") & IIf(i = Neuri, " (" & IIf(üdt.fußst = auff, "Fußst,", "") & IIf(üdt.sens = auff, "Sens,", "") & IIf(üdt.sens = pathdok, "ICD G63.2,", "") & IIf(üdt.puls = auff, "Puls,", "") & IIf(üdt.FEn(12), "<span id='cave'>DFS</span>", "") & "&rArr;" & Fqmin(i) & "/a)", "") & IIf(i = Gwi, " (BMI " & Round(üdt.bmi, 1) & " kg/m²)", "") & IIf(i = Schuli, " (Soll:" & sollz & IIf(obnach, "+", "") & ")", "") & "</th>"
    Else
     If ie < zl Then ie = ie + 1
    End If
    i = i + 1
   Loop
   
   Print #355, "   </TR>"
   For j = 1 To rszmax ' Durchlauf der Kleinzeilen innerhalb eines Blocks
'Werte(j)
    Print #355, "   <TR>"
    For i = ia To ie ' Parameter in jedem Block
     If ((i <> SchulEintri Or Fqmin(Schuli) = 365) And ((nurpath(i) = 0 Or nurpath(i) = 3 Or pathol(i, 0)) And rsz(i) <> 0)) Then
       If j <= rsz(i) * IIf(RowSp(i) = 0, 1, RowSp(i)) Then
        obpath(i) = 0
        If j = 1 Then
         Dim Zp As Date
         Zp = QAnf(ZQuart(rs(i)!Zp))
' wenn z.B. nicht in diesem Quartal (fqmin(i)=4) und auch nicht jünger als die Hälfte des entsprechenden Abstandes, dann rot
         If Fqmin(i) <> 0 Then
          If Now() - Zp > 365 / Fqmin(i) And (Now - rs(i)!Zp) > 365 / Fqmin(i) * 0.5 Then obpath(i) = True
          If obpath(i) Then If i = Fußi Then If obpath(Neuri) Then obpath(Fußi) = False
         End If ' Fqmin(i) <> 0 Then
        End If ' j = 1 Then
        If Pause(i) = 0 Then
         Print #355, "    <td" & IIf(RowSp(i) <> 0, " rowspan='" & min(RowSp(i), rszmax - j + 1) & "' ", "") & " id='schmal'><P" & IIf(obpath(i), " id='cave'", "") & ">" & Format(rs(i)!Zp, "d.m.yy") & "</td>"
         Debug.Print rs(i)!Wert
         Print #355, "    <td" & IIf(RowSp(i) <> 0, " rowspan='" & min(RowSp(i), rszmax - j + 1) & "' ", "") & IIf(pathol(i, j), " id='path'", "") & IIf(Weite(i) <> "", " style='width:" & Weite(i) & "'", "") & ">" & IIf(RowSp(i) <> 0, "", "<p>") & rs(i)!Wert & IIf(RowSp(i) <> 0, "", "</p>") & "</td>"
         If RowSp(i) <> 0 Then
          Pause(i) = RowSp(i) - 1
         End If
         rs(i).MoveNext
        Else
         Pause(i) = Pause(i) - 1
        End If
#Const mitLeerZellen = False
#If mitLeerZellen Then
       ElseIf Pause(i) = 0 Then
'        If RowSp(i) = 0 Or j = 1 Then ' Beim Neurostatus in die folgenden Zeilen nichts mehr schreiben
         Print #355, "    <td rowspan='" & rszmax - j + 1 & "' colspan = '2' id='schmal' style='border-style:none'></td>"
         Pause(i) = -1
       ElseIf Pause(i) > 0 Then
         Pause(i) = Pause(i) - 1
#Else
      Else  ' rsz(i) >= j Then
       If rsz(i) <> 0 Then
'        If i <> Neuri Or j = 1 Then ' Beim Neurostatus in die folgenden Zeilen nichts mehr schreiben
         Print #355, "    <td id='schmal'></td><td></td>"
'        End If ' rsz(i) <> 0 Then
       End If
#End If
       End If ' rsz(i) >= j Then
      End If ' nurpath(i) = 0 Or pathol(i, 0) Then
    Next i
    Print #355, "   </TR>"
    If j = rszmax Then
     If i = zl Then
      Exit For ' fertig
     End If
    End If
   Next j
   Print #355, "</TABLE>"
   If ie = zl Then Exit Do
   k = k + 1
  Loop
  Print #355, "<p id='abstand'> </p>"
  Print #355, "<div id='abstand'>" & Replace$(Replace$(DiagString(CLng(Pid)), vbVerticalTab & " ", "<br />&nbsp;"), vbVerticalTab, "<br />") & "</div>"
  
' Hier kommt das Labor rein!
  Dim SelbstStatus%, raDatBOF%, Matr$(), MForm%(), mBreiten$()
  Call LaborInsPLZ(Pid, SelbstStatus, raDatBOF, Matr, MForm, mBreiten)
  Print #355, "<p id='abstand'> </p>"
  Print #355, "<Table>"
  For j = 0 To UBound(Matr, 2)
   Print #355, " <tr>"
   Dim iumw&
   For i = 3 To UBound(Matr, 1)
    If i > 5 Then
     iumw = UBound(Matr, 1) - i + 6
    Else
     iumw = i
    End If
    Dim Formg$
    Select Case MForm(iumw, j)
     Case 0: Formg = "lnn"
     Case 1: Formg = "lno"
     Case 2: Formg = "lpn"
     Case 3: Formg = "lpo"
    End Select
    Print #355, "<td id='" & Formg & "'>" & Matr(iumw, j) & "</td>"
   Next i
   Print #355, " </tr>"
  Next j
  Print #355, "</Table>"
  Print #355, "</BODY>"
  Print #355, "</HTML>"
  
  Close #355
  DoEvents
  Call SetFileTimeByDate(DateiNameAkt, mftLastWriteTime, Datum + Uhrzeit)
  Lese.Ausgeb "Patientenlaufzettel für " & rnam!Nachname & ", " & rnam!Vorname & " (Pat_id: " & Pid & ") erstellt.", True
  syscmd 4, "Fertig mit dem Patientenlaufzettel für: " & rnam!Nachname & ", " & rnam!Vorname & " (Pat_id: " & Pid & ")"
'  Call Shell("firefox.exe " & DateiName)
'  Dim reg0$, reg1$
 ' reg0 = getReg(HLM, "software\mozilla\Mozilla Firefox", "CurrentVersion")
  If mitanzeig Then Shell (getReg(HLM, "software\mozilla\Mozilla Firefox\" & getReg(HLM, "software\mozilla\Mozilla Firefox", "CurrentVersion") & "\Main", "PathToExe") & " """ & DateiNameAkt & """")
 End If ' not rnam.eof
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dodoPLZ/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dodoPLZ

Sub LaborInsPLZ(Pat_id$, SelbstStatus%, raDatBOF%, ByRef Matr$(), ByRef MForm%(), ByRef mBreiten$()) 'Word.Document, Pat_id&)
 Dim raLw As New ADODB.Recordset, raDat As New ADODB.Recordset, rLaU As New ADODB.Recordset, ls$
' Zeilenzahl bestimmen
 Dim ZZ&, rz$, gschl$, Vgl$, altGruppe%, Nb$ ' Normbereich
 Dim sql1$
 Dim u!, o! ' oberer und unterer Grenzwert numerisch
 Dim uNG$, oNG$ ' obere und untere Normgrenze in Zeichen
 Dim unm$, onm$, unw$, onw$
 Dim pKz% ' Pathologisch-Kennzeichen
 Dim altgru$
 On Error GoTo fehler
 sql = "select Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" as NB from (" & laborAbfr & " where pat_id = " & Pat_id & ") as labor where not (Wert like ""%eine Berechnung%"") " & _
  "UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar,Normbereich as NB " + _
  "FROM laborxus LEFT JOIN laborxwert ON laborxus.RefNr=laborxwert.RefNr " + _
  "WHERE pat_id = " + CStr(Pat_id) + " and not (Wert like ""%eine Berechnung%"") and not exists (select * from laborneu where pat_id = " + CStr(Pat_id) + " and abkü = laborxwert.Abkü and wert = laborxwert.wert and zeitpunkt > laborxus.Eingang-3 and zeitpunkt < laborxus.Eingang+6)"
' sql = "select * from labor1 where pat_id = " & Pat_id & " and not wert like '%eine Berechnung%' order by zeitpunkt desc union select * from labor2 where pat_id = " & Pat_id & " and not wert like '%eine Berechnung%' order by zeitpunkt desc"
 sql = "select * from labor1 where pat_id = " & Pat_id & " and not wert like '%eine Berechnung%' order by zeitpunkt desc union select * from labor2 where pat_id = " & Pat_id & " and not wert like '%eine Berechnung%' order by zeitpunkt desc"
 sql = "select * from labor1 where pat_id = " & Pat_id & " and not wert like '%eine Berechnung%' union select * from labor2 where pat_id = " & Pat_id & " and not wert like '%eine Berechnung%'"
 On Error GoTo fehler
 SelbstStatus = 0
' Set rDat = Dtb.OpenRecordset("Select * from `" + QMdbAkt + "`.labor where pat_id = " & CStr(Pat_id) & " and kommentar = ""Dieser Eintrag wurde manuell erzeugt.""")
 sql1 = "Select * from (" & laborAbfr & " where pat_id = " & Pat_id & ") as labor where kommentar = ""Dieser Eintrag wurde manuell erzeugt."""
 raDat.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
 If Not raDat.BOF Then SelbstStatus = SelbstStatus + 1
' Set rDat = Dtb.OpenRecordset("Select * from `" + QMdbAkt + "`.labor where pat_id = " & CStr(Pat_id) & " and kommentar <> ""Dieser Eintrag wurde manuell erzeugt.""")
 sql1 = "Select * from (" & laborAbfr & " where pat_id = " & Pat_id & ") as labor where kommentar <> ""Dieser Eintrag wurde manuell erzeugt."""
 Set raDat = Nothing
 raDat.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
 If Not raDat.BOF Then SelbstStatus = SelbstStatus + 2
 
 gschl = ""
 altGruppe = 0
' Set rLW = TabÖff("Namen", "Pat_id")
' ralw.Seek "=", Pat_id
 raLw.Open "select geschlecht from namen where pat_id = " & Pat_id, DBCn, adOpenStatic, adLockReadOnly
 If Not raLw.EOF Then
  gschl = raLw!Geschlecht
 End If
 raLw.Close
 
 Dim sql0$
 Dim sql2$
nochmal:
 sql2 = "select * from labor1 where pat_id = " & Pat_id & " union select * from labor2 where pat_id = " & Pat_id
 sql0 = "SELECT distinct gruppe, reihe, sql1.abkü, sql1.einheit from (" + sql2 + ") as sql1 left join laborparameter on sql1.abkü = laborparameter.abkü and iif(isnull(sql1.einheit) or sql1.einheit="""", ""kA"",sql1.einheit) = laborparameter.einheit where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar)) and reihe <> 999 order by gruppe, reihe"
 If lies.obMySQL Then sql0 = Replace$(sql0, "iif(", "if(")
 Set raLw = Nothing
' GoTo nochmal
 raLw.Open sql0, DBCn, adOpenDynamic, adLockReadOnly
 altgru = ""
 Do While Not raLw.EOF
  If raLw!gruppe <> altgru Then
   ZZ = ZZ + 1
   altgru = raLw!gruppe
  End If
  ZZ = ZZ + 1
  raLw.MoveNext
 Loop
 
 syscmd 4, "Labor (1) nach Zeilenzahlbestimmung"
' sql1 = "select count(*) from (SELECT distinct abkü from (" + sql + ") as sql1 where (not isnull(wert) or not isnull(kommentar))) as innen"
'' sql = "select count(*) from (SELECT distinct abkü from `" + QMdbAkt + "`.laborunion where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar)))"
'' Set raLW = Dtb.OpenRecordset(sql1)
' Set raLW = Nothing
' raLW.Open sql1, dbcn, adOpenDynamic, adLockReadOnly
' zZ = raLW.Fields(0)
' raLW.Close
' sql1 = "select count(*) from (SELECT distinct gruppe from (" + sql + ") as sql1 left join laborparameter on sql1.abkü = laborparameter.abkü and iif(isnull(sql1.einheit) or sql1.einheit="""", ""kA"",sql1.einheit) = laborparameter.einheit where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar))) as innen"
'' sql = "select count(*) from (select distinct gruppe from (SELECT * from `" + QMdbAkt + "`.laborunion left join `" + QMdbAkt + "`.laborparameter on laborunion.abkü = laborparameter.abkü and iif(isnull(laborunion.einheit) or laborunion.einheit="""", ""kA"",laborunion.einheit) = laborparameter.einheit where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar))))"
'' Set rLW = Dtb.OpenRecordset(sql1)
' If lies.obmysql Then sql1 = replace$(sql1, "iif(", "if(")
' Set raLW = Nothing
' raLW.Open sql1, dbcn, adOpenDynamic, adLockReadOnly
' zZ = zZ + raLW.Fields(0)
' raLW.Close
nochmal1:
 sql1 = "select distinct * from (SELECT distinct datevalue(zeitpunkt) as Datum from (" & sql & ") as sql1 left join laborparameter on sql1.abkü = laborparameter.abkü and iif(isnull(sql1.einheit) or sql1.einheit="""", ""kA"",sql1.einheit) = laborparameter.einheit where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar)) and reihe <> 999) as i0 order by datum"
' das erste distinct ist für Access nötig
' Set rDat = Dtb.OpenRecordset("select count(*) from (" + sql1 + ")")
' rZ = radat.Fields(0) ' dateserial(year(zeitpunkt),month(zeitpunkt),day(zeitpunkt))
 Set raDat = Nothing
 If lies.obMySQL Then
  sql1 = Replace$(Replace$(sql1, "datevalue(", "date("), "iif(", "if(")
 End If
' GoTo nochmal1
 raDat.Open "select count(*) as ct from (" + sql1 + ") as innen", DBCn, adOpenDynamic, adLockReadOnly
 If Not raDat.BOF Then
  rz = raDat!ct
 End If
 
 raDat.Close
 ReDim Matr(rz + 5, ZZ + 1) ' Spaltenzahl, Zeilenzahl, matr(x,y)= Wert
 ReDim MForm(rz + 5, ZZ + 1) ' Formatierung (1 = schräg, 2 = fett)
 ' links noch: 0:Abkü, 1:unterer NW, 2:oberer NW, dann 3:Langtext, 4:Einheit, 5:Normbereich
 ReDim mBreiten(rz)
 
 
' Tabl.cell(1, 1).Range = "Parameter"
' Tabl.cell(1, 2).Range = "Einheit"
' Tabl.cell(1, 3).Range = "Normbereich"
 Matr(3, 0) = "Parameter"
 Matr(4, 0) = "Einheit"
 Matr(5, 0) = "Normbereich"
 Set raDat = Nothing
 raDat.Open sql1, DBCn, adOpenDynamic, adLockReadOnly
' Set rDat = Dtb.OpenRecordset(sql1)
 If raDat.BOF Then raDatBOF = True
 syscmd 4, "Labor (1) nach Überschrift"
 
 Dim i&, j&
 i = 1
' raDat.Move DiffBr

 Do While Not raDat.EOF
  Matr(i + 5, 0) = Format$(raDat!Datum, "dd.mm.yy")
  i = i + 1
  raDat.Move 1
  If i > rz Then Exit Do
 Loop
 
 syscmd 4, "Labor (1) nach Datumseintragung"
 
 Dim rLP As New ADODB.Recordset
 j = 1
 altgru = ""
 raLw.MoveFirst
 Do While Not raLw.EOF
  If raLw!gruppe <> altgru Then
   altgru = raLw!gruppe
   Matr(0, j) = "Gruppe " & raLw!gruppe & "|" & raLw!reihe
'   Debug.Print Matr(0, j) & "," & Matr(1, j) & "," & Matr(2, j)
   j = j + 1
  End If
  Matr(0, j) = raLw!Abkü
  Matr(4, j) = raLw!Einheit
  Set rLP = Nothing
  Dim rLpgeht%
  rLpgeht = -1
  rLP.Open "select * from laborparameter where abkü = '" & raLw!Abkü & "' and einheit = '" & raLw!Einheit & "' and (unm <> """" or unw <> """" or onm <> """" or onw <> """")", DBCn, adOpenStatic, adLockReadOnly
  If rLP.BOF Then rLpgeht = 0
  If rLpgeht Then If IsNull(rLP!Langtext) Then rLpgeht = 0
  If rLpgeht Then If rLP!Langtext = "" Then rLpgeht = 0
  If rLpgeht Then
    Matr(3, j) = rLP!Langtext
  Else
   Dim rLangt As New ADODB.Recordset
   Set rLangt = Nothing
   rLangt.Open "select LangText from (" & laborAbfr & " where abkü = '" & raLw!Abkü & "' and einheit = '" & raLw!Einheit & "' and not isnull(langtext) and langtext <> """") as labor", DBCn, adOpenStatic, adLockReadOnly
   If rLangt.BOF Then
    Set rLangt = Nothing
    rLangt.Open "select LangText from (" & laborAbfr & " where abkü = '" & raLw!Abkü & "' and not isnull(langtext) and langtext <> """") as labor", DBCn, adOpenStatic, adLockReadOnly
   End If
   If Not rLangt.BOF Then
    Matr(3, j) = rLangt!Langtext
   End If
  End If
  If Not rLP.EOF Then
   If gschl = "m" Or (rLP!unw = "" And rLP!onw = "") Then
    Matr(5, j) = rLP!unm & "-" & rLP!onm
    Matr(1, j) = Replace$(nz(rLP!unm,""), "1:", "")
    Matr(2, j) = Replace$(nz(rLP!onm,""), "1:", "")
   Else
    Matr(5, j) = rLP!unw & "-" & rLP!onw
    Matr(1, j) = Replace$(nz(rLP!unw,""), "1:", "")
    Matr(2, j) = Replace$(nz(rLP!onw,""), "1:", "")
   End If
   If Matr(0, j) Like "LDL*" Then Matr(5, j) = "-100"
  End If
'  Debug.Print Matr(0, j) & "," & Matr(1, j) & "," & Matr(2, j)
  j = j + 1
  raLw.Move 1
 Loop
 
 syscmd 4, "Labor (1) nach Normbereichseingabe"
 
 Set raLw = Nothing
 raLw.Open sql, DBCn, adOpenDynamic, adLockReadOnly
 Do While Not raLw.EOF
  i = 6
  Do While CDate(Matr(i, 0)) < DateValue(raLw!Zeitpunkt)
   i = i + 1
   If i >= rz + 5 Then Exit Do
  Loop
  j = 0
  If i > UBound(Matr, 1) Then
   i = UBound(Matr, 1)
   MsgBox "Stop in LaborIns: " & vbCrLf & "i: " & i & vbCrLf & "ubound(matr,1): " & UBound(Matr, 1)
   Stop
  End If
  Do While j <= UBound(Matr, 2)
   If Matr(0, j) = raLw!Abkü Then
    If Matr(4, j) = raLw!Einheit Then
     Exit Do
    End If
   End If
   j = j + 1
  Loop
  If j > UBound(Matr, 2) Then j = UBound(Matr, 2)
  Matr(i, j) = raLw!Wert
  If Not IsNull(raLw!Kommentar) And raLw!Kommentar <> "" Then
   If InStrB(raLw!Kommentar, "manuell") <> 0 Then
'     Tabl.cell(j + 1, i - 2).Range.Font.Italic = True
     MForm(i, j) = MForm(i, j) + 1
   ElseIf raLw!Wert = "" Then
'     Matr(i + Diffbr, j) = Matr(i, j) & IIf(Matr(i, j) = "", "", " ") & raLW!Kommentar
     Matr(i, j) = raLw!Kommentar
   End If
  End If
  pKz = 0 ' "n" normal
  If IsNumeric(raLw!Wert) Or raLw!Wert Like "1:*" Or raLw!Wert Like "<*" Or raLw!Wert Like ">*" Then
   Dim rlWS$, rlWZ#, obkl%, obgr%
   Dim buch$, pos&
   obkl = 0
   obgr = 0
   rlWS = raLw!Wert
   For pos = 1 To Len(raLw!Wert)
    If InStrB("0123456789,.:-+<>", Mid$(rlWS, pos, 1)) = 0 Then
     rlWS = Left(rlWS, pos - 1) & Mid$(rlWS, pos + 1)
     pos = pos - 1
    End If
   Next pos
   If rlWS Like "1:<*" Then rlWS = "<1:" & Mid$(rlWS, 4)
   If Left(rlWS, 1) = "<" Then
    obkl = True
    rlWS = Mid$(rlWS, 2)
   ElseIf Left(rlWS, 1) = ">" Then
    obgr = True
    rlWS = Mid$(rlWS, 2)
   End If
   rlWS = Replace$(Replace$(Replace$(Replace$(rlWS, "^", ""), "1:", ""), ".", ","), ",,", ",")
   If rlWS = "" Then rlWS = "0"
   rlWZ = CDbl(rlWS)
   If LenB(Matr(1, j)) <> 0 Then
    If rlWZ < CDbl(Replace$(Matr(1, j), ".", ",")) And Not obgr Then
     pKz = 1 ' "p"
    End If
   End If
   If LenB(Matr(2, j)) <> 0 Then
    On Error Resume Next
    If rlWZ > CDbl(Replace$(Matr(2, j), ".", ",")) And Not obkl Then
     pKz = 1 ' "p"
    End If
    On Error GoTo fehler
   End If
'   If o <> -99 And rlwd > o Then pKz = "p"
  End If
  On Error Resume Next
  If pKz = 1 Then '"p" Then 'If Tabl.Cell(j + 1, i - 2) <> "" Then
'   Tabl.cell(j + 1, i - 2).Range.Font.Bold = True
   MForm(i, j) = MForm(i, j) + 2
  End If
  On Error GoTo fehler
'  If IsNull(raLW!gruppe) Then
'   altgruppe = 0
'  Else
'   altgruppe = Val(raLW!gruppe)
'  End If
  raLw.MoveNext
 Loop
' dc.Bookmarks!laborTab.Range
 'sql1 = "select * from (SELECT distinct(cdate(int(zeitpunkt))) as Datum from (" + sql + ") as sql left join `" + QMdbAkt + "`.laborparameter on sql.abkü = laborparameter.abkü and iif(isnull(sql.einheit) or sql.einheit="""", ""kA"",sql.einheit) = laborparameter.einheit where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar))) order by datum"
' Set rLW = Dtb.OpenRecordset(sql1, dbOpenDynaset)
 
 For j = 0 To ZZ
  If Matr(3, j) = "HbA1c Eigenlabor" Then
   Matr(3, j) = "HbA1c"
  End If
 Next j
 syscmd 4, "Labor (1) nach Dateneintragung "

 
 ' ls = CStr(rZ) & vbcrlf
' raDat.MoveFirst
' Dim CNr%
' CNr = rZ + 3
' raDat.MoveLast
' Do While Not raDat.BOF
'  ' ls = LS + format$(radat!Datum, "dd.mm.yyyy") & vbcrlf
'  If lies.obmysql Then
'   Tabl.cell(1, CNr).Range = raDat!Datum 'BDTtoDate(mid$(raDat!Datum, 7, 2) & mid$(raDat!Datum, 5, 2) & left(raDat!Datum, 4))
'  Else
'   Tabl.cell(1, CNr).Range = format$(raDat!Datum, "dd.mm.yy")
'  End If
'  CNr = CNr - 1
'  raDat.Move -1
' Loop

' ' ls = LS + CStr(zZ) & vbcrlf
' sql1 = "select * from (SELECT reihe, gruppe, unw, onw, unm, onm, sql1.*, datevalue(zeitpunkt) as Datum, sql1.abkü as sabkü, sql1.einheit as seinheit, sql1.langtext as slangtext from (" + sql + ") as sql1 left join laborparameter on sql1.abkü = laborparameter.abkü and iif(isnull(sql1.einheit) or sql1.einheit="""", ""kA"",sql1.einheit) = nz(laborparameter.einheit,"""") where pat_id = " + CStr(Pat_id) + " and (not isnull(wert) or not isnull(kommentar))) as i0 order by gruppe,reihe,datum" ' dateserial(year(zeitpunkt),month(zeitpunkt),day(zeitpunkt))
' If lies.obmysql Then sql1 = replace$(replace$(sql1, "datevalue(", "date("), "iif(", "if(")
' Set raLW = Nothing
'' Set rLW = Dtb.OpenRecordset(sql1)
' raLW.Open sql1, dbcn, adOpenDynamic, adLockReadOnly
' Dim aktZ%, VorWert$, AktGru$, AktSp%, AktAbkü$, obgeändert%
' Do While Not raLW.EOF
'  If Not IsNull(raLW!sabkü) Then
'   If IsNull(raLW!gruppe) Then
''    If rLaU Is Nothing Then Set rLaU = Dtb.OpenRecordset("select * from (" + sql + ") as sql order by zeitpunkt")
'    Set raLau = Nothing
'    raLau.Open "select *,sql.abkü as sabkü, sql.langtext as slangtext from (" & sql & ") as sql order by zeitpunkt", dbcn, adOpenDynamic, adLockReadOnly
'    Debug.Print "Kein Parametereintrag für: " + CStr(raLW!sabkü) & " " & CStr(nz(raLW!slangtext,"")) + " `" + nz(raLW!seinheit,"") + "`"
'    Call raLau.Find("Pat_id = " & CStr(raLW!Pat_id) & " and zeitpunkt = cdate(""" & CStr(raLW!Zeitpunkt) & """) and fertigstgrad = """ & CStr(raLW!FertigStGrad) & """ and abkü = """ & CStr(raLW!sabkü) & """", 0, adSearchForward, 1)
    
'    If Not rLaU.BOF Then
'     Call LaborParameter(raLau)
'     obgeändert = True
'    End If
'   End If
'  End If
'  raLW.Move 1
' Loop
' If obgeändert Then
'  Set raLW = Nothing
''  Set raLW = Dtb.OpenRecordset(sql)
'  raLW.Open sql, dbcn, adOpenDynamic, adLockReadOnly
' Else
' raLW.MoveFirst
' End If
' aktZ = 1
'  Dim TestFeld$
'  Dim TestInh
'  On Error Resume Next
'  Err.Clear
'  TestFeld = "sql."
'  TestInh = raLW.Fields(TestFeld + "abkü")
'  If Err.Number <> 0 Then
'   TestFeld = ""
'   Err.Clear
'  End If
'  On Error GoTo fehler
' Do While Not raLW.EOF
'  If Not IsNull(raLW.Fields(TestFeld + "abkü")) Then
''   If Val(ralw!Gruppe) <> altGruppe Then
'    ' ls = LS & vbcrlf
''   End If
'   If IsNull(VorWert) Or VorWert = "" Or raLW.Fields(TestFeld + "LangText") <> VorWert Then
'    aktZ = aktZ + 1
'    AktSp = 4
'    On Error GoTo gruppenfehler
'    If Not IsNull(AktGru) And AktGru <> "" And raLW.Fields("Gruppe") <> AktGru Then
'    On Error GoTo fehler
'      aktZ = aktZ + 1
'    End If
'    Do While aktZ > Tabl.Rows.Count
'     Tabl.Rows.Add
'     zZ = zZ + 1
'    Loop
'    AktGru = nz(raLW.Fields("Gruppe"),"")
'    Tabl.cell(aktZ, 1).Range = raLW.Fields("slangtext") 'nz(raLW.Fields("sql.LangText"),"")
'    Tabl.cell(aktZ, 2).Range = raLW.Fields("seinheit") 'nz(raLW.Fields("sql.Einheit"),"")
'    u = -99
'    o = -99
'    uNG = ""
'    oNG = ""
'    unw = nz(raLW!unw,"")
'    onw = nz(raLW!onw,"")
'    unm = nz(raLW!unm,"")
'    onm = nz(raLW!onm,"")
''    If instrb(raLW.Fields("sql.Langtext"), "LDL") > 0 Then
'    If instrb(raLW.Fields("slangtext"), "LDL") > 0 Then
'     onm = "100"
'    End If
'    If gschl = "w" Then
'     uNG = replace$(unw, "1:", "")
'     oNG = replace$(onw, "1:", "")
'    End If
'    If uNG = "" And oNG = "" Then
'     uNG = replace$(unm, "1:", "")
'     oNG = replace$(onm, "1:", "")
'    End If
'    If uNG <> "" Then u = CDbl(uNG)
'    If oNG <> "" Then o = CDbl(oNG)
''   If IsNumeric(raLw!Wert) Or raLw!Wert Like "1:*" Then
''    If gschl = "w" Then
''     If Not ralw!unw = "" Then u = CDbl(replace$(ralw!unw, "1:", ""))
''     If Not ralw!onw = "" Then o = CDbl(replace$(ralw!onw, "1:", ""))
''    End If
''    If gschl = "m" Or u = -99 Then
''     If Not ralw!unm = "" Then u = CDbl(replace$(ralw!unm, "1:", ""))
''    End If
''    If gschl = "m" Or o = -99 Then
''     If Not ralw!onm = "" Then o = CDbl(replace$(ralw!onm, "1:", ""))
''    End If
''   End If
'    ' ls = LS + pKz & vbcrlf
'    NB = ""
'    If uNG <> "" Or oNG <> "" Then
'     If gschl = "w" And (unw <> "" Or onw <> "") Then
'      NB = unw + "-" + onw
'     End If
'    End If
'    If NB = "" And (unm <> "" Or onm <> "") Then
'     NB = unm + "-" + onm
'    End If
'    Tabl.cell(aktZ, 3) = NB
'    ' ls = LS + NB & vbcrlf
'   End If
''   VorWert = nz(raLW.Fields("sql.LangText"),"")
'   VorWert = raLW.Fields("slangtext") 'nz(raLW.Fields("sql.LangText"),"")
'   ' ls = LS + ralw.Fields("sql.LangText") & vbcrlf
'   ' ls = LS + ralw.Fields("sql.Einheit") & vbcrlf
'   ' ls = LS + format$(ralw!Datum, "dd.mm.yyyy") & vbcrlf
'   ' ls = LS + ralw!Wert & vbcrlf
   
'   Do
'    If DateValue(raLW!Datum) > CDate(left(Tabl.cell(1, AktSp).Range, 8)) Then
'     AktSp = AktSp + 1
'    Else
'     Exit Do
'    End If
'    If AktSp > rZ + 3 Then
'     Exit Do
'    End If
'   Loop
   
'   Tabl.cell(aktZ, AktSp).Range = IIf(IsNull(raLW!Wert), IIf(IsNull(raLW!Kommentar), "", format$(raLW!Zeitpunkt, "dd/mm/yy") + ": " + raLW!Kommentar), replace$(replace$(trim$(nz(raLW!Wert,"")), vbcr, ""), vbtab, ""))
'   If instrb(raLW!Kommentar, "manuell") > 0 Then Tabl.cell(aktZ, AktSp).Range.Font.Italic = True
   
'  End If ' not isnull(ralw!`sql!abkü`
'  raLW.Move 1
' Loop '  While Not ralw.EOF
 
' Open "c:\wwtest.txt" For Output As #34
' Print #34, LS
' Close #34
 
 raDat.Close
 raLw.Close
  
 Exit Sub
'gruppenfehler:
'    If Err.Number = 3265 Then ' Gruppe fehlt in Auflistung, 18.4.06
'     MsgBox "Resume in LaborIns: Err.Nummer = 3265"
'     Resume
'     Resume Next
'    End If
'    Exit Sub
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.path
#End If
Select Case MsgBox("FNr: " & FNr & "ErrNr: " & CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.source), vbNullString, CStr(Err.source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborInsPLZ/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): Ende
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Sub      ' LaborIns

Function doPLZeinzeln()
 Dim erg$
 erg = InputBox("Bitte Pat-ID aus Turbomed eingeben:", "Patientenlaufzettel Eingabe")
 If IsNumeric(erg) Then
  Call dodoPLZ(erg, Now, Now - Int(Now), True)
 End If
End Function
