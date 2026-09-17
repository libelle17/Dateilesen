Attribute VB_Name = "Importiere"
Option Explicit
'Const obMy% = True ' ob ADO-Quelle eine MySQL-Datenbank ist
Dim numA&
Dim rs As ADODB.Recordset
Dim rEzä() As ADODB.Recordset
'Public rAna As ADODB.Recordset
Dim obvorgestellt As Boolean, messDatum As Date
Dim FormID&
Dim FormAbk$, FormBez$, FormVorl$, FormInh$
Public FoIDv& ' Pseudo-Foid
Dim DokPfad$, DokArt$, DokName$
'Dim lVorl As Date ' letzte Vorlage, 4109
Public VorByte&, AktByte&, rEzäBeg As Date
Public Const qverz$ = "u:\"
Public Const aVerz$ = qverz + "Anamnese\"
Public Const eVerz$ = "P:\Patientenübergreifendes"
Public Const üVerz$ = qverz + "TMExport\"
Public Arra() ' Array für Telefonnummern
Public ArraInd%
Dim lfdfl&, AktZeit As Date, f8000$, f8100$
Public LAktZeit As Date ' bei Namen steht die letzte Datenänderung aus irgend einem zu dem Namen gehörigen Satz aus BDT-Datei drin, um nur aktuelle aktualisieren zu können
Dim Diag$(), ICD$(), DiagSi$(), DiagSe$(), DiagAttr$(), diagDat() As Date, obDauer%()
Dim gk() As Boolean
Dim DiagNr%
Dim obkNeph% ' ob keine Nephropathie zu bestehen scheint (-1 => vorliegende Diagnose wird ignoriert)
'Dim nachname$, vorname$, GebDat As Date, Titel$, nvorsatz$, GesamtName$

Const DiagMaxZahl% = 40
' Variablen für Anamnesebogen
Dim DMSchulz%, DMSchL%, RRSchulz%, DMPDat As Date, Tkz%, VorStDat As Date, obMedNetz%
Dim medplan$, obMedPlan%, obMedPlanGelesen%
Dim NEin0% ' Nameneintr0 wurde aufgerufen
Dim Arr$, gesRR$, lRR$  ' gesamter, Anfangs- und letzter Blutdruck
Dim obEinweisung%
Dim EintrArt$, EintrInh$
Dim FStG$, Abkü$, Langtext$, Wert$, Einheit$, Kommentar$
Dim AlbErhöht$, AlbDat As Date
Dim lHbA1c$, lHbA1cDat As Date  ' für Anamnesebogen
Dim absPos&
Dim DStr$ ' Datumsumwandel-String für Blutdruckangabe
Dim MPDatum As Date
'Public Med$(), Dos$()
Public MedNr%, MedZahl%, MPNr& ' für Medikamentenplan
Dim MedZeit%, obMed As Boolean, lDos%, MedArt$
Dim ZeileNr% ' wenn mehrere 6299 auf ein 6298 folgen, dann können alle eingetragen werden
Dim raf& ' records affected
Public rUn1&, rFo1&, rFi1& ' erster neuer Datensatz in formulare / forminhaltform_abk
Public Const obForeign% = True
Public obMitAlterTab% ' alter table bei zu kurzen Feldern in MySQL
Public PcDokPfad$
Public Fil As File
Public MedSL As SortierListe ' für MedKlass
 ' ehemals statische Variablen einzelner Prozeduren, hier für erneuten Einlesvorgang rausgezogen
Public rAnaAdo As New ADODB.Recordset
Dim Quartal$ ' 4101
Dim lfdnr&
Dim lKennung% ', lk%(15)
Dim jetztKopf%, rFm_Nr&
Dim lFormID& ' ID des letzten Formulars
Dim obZweiteRunde%
Dim lLang$, lLangVW&
Dim lKomm$, lKommVW&
Dim SpeicherZt As Date
Dim aktQ$ ' aktuelles Quartal
Dim obDStr% ' ob Diagnose in Diagnosestring einzutragen ist

'Public rAna() As Ana
'Public rFm() As forminhfeld
Public ls% 'Laborsatz
Public sFld As sFeld, sListFeld As New SortierListe
Public sFldI As sFeldInh, sListFldInh As New SortierListe
Public sLp As sLpar, sListLpar As New SortierListe
Function doTabVorb(frm As Lese, obInhalt%)
 Dim i&, t2#
 On Error GoTo fehler
 Dim rs As New ADODB.Recordset, raf&
' Dim cn As New ADODB.Connection,
' cn.Open CSStr, , , 0
' Set rs = cn.Execute("select * from quelle.forminhaltfeld", rAf) ' + adAsyncExecute)
' Call rs.Open("quelle.forminhaltfeld", cn, adOpenStatic, adLockReadOnly)
 Set rs = DBCn.Execute("select abkü, langtext, einheit from laborparameter")
 Do While Not rs.EOF
  Set sLp = New sLpar
  sLp.Abkü = nz(rs!Abkü,"")
  sLp.Einheit = nz(rs!Einheit,"")
  sLp.Langtext = nz(rs!Langtext,"")
  Call sListLpar.sCAdd(sLp, -1)
  rs.Move 1
 Loop
 rs.Close
 Set rs = DBCn.Execute("select feld,feldvw from forminhaltfeld")
 Do While Not rs.EOF
  Set sFld = New sFeld
  sFld.Feld = nz(rs!Feld,"")
  sFld.FeldVW = rs!FeldVW
' If Not sListFeld.SuchItem(sFld) Then
  Call sListFeld.sCAdd(sFld, -1)
' End If
  rs.Move 1
 Loop
' For i = 1 To sListFeld.Count
'  Debug.Print sListFeld.Item(i).Feld
' Next i
 rs.Close
 If obInhalt Then
 ' Call rs.Open("quelle.forminhaltfeldinh", cn, adOpenStatic, adLockReadOnly)
  Set rs = DBCn.Execute("select feldinh,feldinhvw from forminhaltfeldinh", raf) ', adAsyncExecute)
  frm.Bytes = "(FormInhaltFeldInh)"
  Do While Not rs.EOF
   Set sFldI = New sFeldInh
   sFldI.FeldInh = nz(rs!FeldInh,"")
   sFldI.FeldInhVW = rs!FeldInhVW
 '  If Not sListFldInh.SuchItem(sFldI) Then
    Call sListFldInh.sCAdd(sFldI, -1)
 '  End If
   rs.Move 1
   i = i + 1
   If i Mod 100 = 0 Then
    frm.Zeilen = i
    DoEvents
   End If
  Loop
  frm.Bytes = 0
  t2 = Now
 ' For i = 1 To sListFldInh.Count
 '  Debug.Print sListFldInh.Item(i).FeldInh
 ' Next i
 End If ' obInhalt
 
 Set rs = DBCn.Execute("select kennung from " & kla & "unbekannte kennungen" & klz, raf)
 Do While Not rs.EOF
  ReDim Preserve rUn(UBound(rUn) + 1)
  rUn(UBound(rUn)).Kennung = rs!Kennung
  rs.Move 1
 Loop
 rUn1 = UBound(rUn)
 
 rFo(0).FormID = 0
' Call rAdo.Open("select * from formulare order by formid", dbcn, adOpenDynamic, adLockReadOnly)
 Set rs = DBCn.Execute("select * from formulare order by formid")
 If Not rs.BOF Then
  rs.MoveFirst
  Do While Not rs.EOF
   ReDim Preserve rFo(UBound(rFo) + 1)
   rFo(UBound(rFo)).Form_Abk = nz(rs!Form_Abk,"")
   rFo(UBound(rFo)).FormBez = rs!FormBez
   rFo(UBound(rFo)).FormVorl = rs!FormVorl
   rFo(UBound(rFo)).FormID = rs!FormID
   rs.MoveNext
  Loop
 End If
 rFo1 = UBound(rFo)
 
 rFi(0).Form_AbkVW = 0
 Set rs = DBCn.Execute("select * from forminhaltform_abk order by form_abkvw")
 If Not rs.BOF Then
  rs.MoveFirst
  Do While Not rs.EOF
   ReDim Preserve rFi(UBound(rFi) + 1)
   rFi(UBound(rFi)).Form_Abk = nz(rs!Form_Abk,"")
   rFi(UBound(rFi)).Form_AbkVW = rs!Form_AbkVW
   rs.MoveNext
  Loop
 End If
 rFi1 = UBound(rFi)
 
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in doFormVorb/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' doFormVorb
Function Zinit()
 If obMySQL Then
  kla = "`"
  klz = "`"
 Else
  kla = "["
  klz = "]"
 End If
End Function ' Zinit
Function löschBezügeausLaborux(Optional RefNr&)
   Call DBCn.Execute("update laborxus set " & _
                     "ZeitpunktLaborneu = " & datForm(CDate(0)) & "," & _
                     "verglichen = null," & _
                     "LWerte = """"," & _
                     "ZdüP = 0," & _
                     "AfN = 0," & _
                     "Pat_idLaborNeu = """" " & _
                     IIf(RefNr <> 0, "where refnr = " & RefNr, ""), raf)
   Call DBCn.Execute("update laborxus set " & _
                     "pat_idursp = ""E"" " & _
                     IIf(RefNr <> 0, "where refnr = " & RefNr & " and", "where") & " pat_idursp in (""B"",""W"")", raf)
   Call DBCn.Execute("update laborxus set " & _
                     "pat_idursp = """"," & _
                     "pat_id = 0 " & _
                     IIf(RefNr <> 0, "where refnr = " & RefNr & " and", "where") & " pat_idursp = ""L""", raf)
End Function ' löschBezügeausLaborux
' Versuch mit Load Data Infile sehr negativ ausgefallen: 45" ggü. 9" Einlesezeit und falscher Zeichensatz,
' unabhängig vom Zeichensatz der der Datenbank (auch Dos usw. versucht)
' falls nicht über MySQL Query Browser eingelesen!
' Datenbank einl2 daher wieder gelöscht. Hier noch die Creates:
'
'`einl2`.CREATE DATABASE `einl2` /*!40100 DEFAULT CHARACTER SET cp850 */;
'
'CREATE TABLE  `einl2`.`latin2` (
'  `lenge` varchar(3) DEFAULT NULL,
'  `kennung` varchar(4) DEFAULT NULL,
'  `inhalt` varchar(150) DEFAULT NULL
') ENGINE=MyISAM DEFAULT CHARSET=latin2 ROW_FORMAT=DYNAMIC;
'
'CREATE TABLE  `einl2`.`zuleing` (
'  `dname` varchar(120) DEFAULT NULL,
'  `bytes` varchar(12) DEFAULT NULL,
'  `geändert` datetime DEFAULT NULL
') ENGINE=MyISAM DEFAULT CHARSET=cp850;
Function GesLies(frm As Lese, Datei$, Optional ZS As Boolean, Optional t1 As Date, Optional LDInF%)
 Dim rAdo As New ADODB.Recordset
 Dim txt$, ltxt$ 'letzter Text
 Dim RLng$, RKennung$, RInhalt$
 Dim Cpt$, t1a As Date, i&
 On Error GoTo fehler
 
 Call StatischInit
 Open üVerz + "ÜbertrProtMySQL.txt" For Output As #322 ' Nummer ab 256: gleichzeitig für andere Prozesse zugreifbar
 If LDInF = 0 Then
  Open Datei For Input As #324
 End If
 Call Tinit
 Call Zinit
#If False Then
 If ProgrammLauf(0, Cpt) Then ' wenn Programm schon (anderweitig) aufgerufen, dann nicht nochmal gleichzeitig aufrufen
  MsgBox "Programm schon auf " + Cpt + " aufgerufen, breche ab!"
  GoTo abbrech
 End If
#End If
 If AllePat Then
  Call AllesLösch(frm)
'  Call löschBezügeausLaborux
 End If
 For i = 0 To frm.Controls.Count - 1
  If frm.Controls(i).Name = "Beenden" Or frm.Controls(i).Name = "Zurücksetzen" Then
   frm.Controls(i).Enabled = False
  ElseIf frm.Controls(i).Name = "Abbrechen" Then
   frm.Controls(i).Enabled = True
  End If
 Next i
 
 Call doTabVorb(frm, obVorb)
 Set MedSL = New SortierListe ' für MedKlass
 Call EintragSpäter(1)
 If LDInF <> 0 Then
  If LDInF = -1 Then
   Call DBCn.Execute("delete from einl2.latin1")
   Call DBCn.Execute("load data infile '" & Replace(frm.QDatei, "\\linux\daten\", "/DAT/") & "' into table einl2.latin1 fields terminated by '' lines terminated by '\r\n';")
   Call DBCn.Execute("delete from einl2.zuleing")
   Dim altobmysql%
   altobmysql = obMySQL
   obMySQL = True
   Call DBCn.Execute("insert into einl2.zuleing (dname,bytes,geändert) values('" & frm.QDatei & "'," & Replace(frm.GesBytes, ".", "") & "," & datForm(frm.QDatum) & ")")
   obMySQL = altobmysql
  Else
   If Not rAdo Is Nothing Then If rAdo.State = 1 Then rAdo.Close
   Call rAdo.Open("einl2.zuleing", DBCn, adOpenStatic, adLockReadOnly)
   frm.QDatei = rAdo!dname
   frm.QDatum = rAdo!geändert
   frm.GesBytes = rAdo!Bytes
  End If
 End If
 If obMySQL And obForeign Then Call DBCn.Execute("set foreign_key_checks = 0")
 If LDInF <> 0 Then
  If Not rAdo Is Nothing Then If rAdo.State = 1 Then rAdo.Close
  rAdo.CursorLocation = adUseServer ' adUseClient 110" für einen Durchlauf adUseServer 73", ob dynamisch oder statisch spielt keine Rolle
'  t1 = Now
''  Set rAdo = DBCn.Execute("select * from einl2.latin1") ', DBCn, adOpenDynamic, adLockReadOnly / Execute 70" ggü open mit 73"
''  DBCn.Close
''  DBCn.Open CStrMy & "einl2" & ";option=" & opti2 ' mit opti1 und opti2 nur 45s
  rAdo.Open "select * from einl2.latin1", DBCn, adOpenDynamic, adLockReadOnly
'  Do While Not rAdo.EOF
'   rAdo.Move 1
'  Loop
'  t1a = Now
'  MsgBox (t1a - t1) * 60 * 60 * 24
 End If
 
 t1a = Now
 Dim obEOF%, obAngeh%, obSchluss%
 If LDInF <> 0 Then
  obEOF = rAdo.EOF
 Else
  obEOF = EOF(324)
 End If
 Dim cKenn(1000000) As String * 4
 Dim cInha(1000000) As String * 72 ' feste Breiten 12", variable 14"; 100000: 11"; 300000: 12", 500000: 12",1000000: 12"
 Dim Ind&, altPat$, aktDat As Date
 Do
  If obEOF Then Exit Do
  If LDInF = 0 Then
   t1 = Now
   Do
    Line Input #324, txt
    RLng = Left(txt, 3)
    RKennung = Mid(txt, 4, 4)
    RInhalt = Mid(txt, 8)
    Select Case RKennung
     Case 3000
      If RKennung = 3000 And RInhalt <> altPat Then
       Ind = 0
       altPat = RInhalt
       LAktZeit = 0
      End If
'     Case 3610, 3649, 5999, 4102, 4109, 4150, 5000, 6200 ' dauert 10" länger als die kurze Version!
     Case 5000, 6200
      aktDat = BDTtoDate(RInhalt)
      If aktDat > LAktZeit Then LAktZeit = aktDat
'     Case 4110, 6201
     Case 6201
      aktDat = aktDat + BDTtoTime(RInhalt)
      If aktDat > LAktZeit Then LAktZeit = aktDat
     ' 4110: Uhrzeit letzter Vorlage zu 4109
     ' 6201: zu 5000
    End Select
    cKenn(Ind) = RKennung
    cInha(Ind) = RInhalt
    Ind = Ind + 1
   Loop While Not EOF(324)
   t1a = Now
   MsgBox (t1a - t1) * 60 * 60 * 24 ' dauert hingegen nur 9" für die ganze Datei und hat den richtigen Zeichensatz, mit Berechnung von LAktZeit 32"
  Else
  End If
  
   frm.Zeilen = frm.Zeilen + 1
   If LDInF = 0 Then
    frm.Bytes = Format(frm.Bytes + Len(txt) + 2, "###,###,###,###,###,###,###,##0")
   Else
    frm.Bytes = Format(frm.Bytes + Len(rAdo!Inhalt) + 7 + 2, "###,###,###,###,###,###,###,##0")
   End If
   If frm.Zeilen Mod 10 = 0 Then
    frm.Sekunden = Format((Now - t1) * 60 * 60 * 24, "###,###,###,###,###,###,##0")
    If frm.GesBytes <> "" Then frm.Prozent = Format(frm.Bytes / frm.GesBytes * 100, "0.00")
    frm.Beginn = Format(t1, "hh:mm:ss")
    If frm.GesBytes <> "" Then
     frm.GesDauer = Format((t1a - t1 + frm.GesBytes / frm.Bytes * (Now - t1a)), "hh:mm:ss")
     frm.Ende = Format(t1 + (t1a - t1 + frm.GesBytes / frm.Bytes * (Now - t1a)), "hh:mm:ss")
    End If
    DoEvents
   End If
  If LDInF <> 0 Then
   obEOF = rAdo.EOF
  Else
   obEOF = EOF(324)
  End If
' wenn LDInF = 0, wird die BDT-Datei direkt gelesen
' wenn LDInF = -1, wird die BDT-Datei zunächst über Load Infile in Tabelle geladen
' wenn LDInf = 2, wird die Tabelle verwendet, ohne dass die BDT-Datei neu geladen wird
' wenn BDT-Datei direkt verwendet wird, so wird der aktuelle Satz zuerst nach txt eingelesen
'                Ist es ein neuer Satz, so wird dann ltxt verwendet, durch txt ersetzt und dann von vorne
'                Ist es ein Anhängsel, so wird es zunächst angehängt, dann eine neue Zeile gelesen usw.
' wenn Tabelle verwendet wird, dann dient RInhalt etwa als ltxt und rado!Inhalt als txt (beim Anhängen auch rado!rlng und rado!kennung
  If LDInF = 0 Then
   If Not Left(txt, 7) Like "#######" Then 'IsNumeric(Left(txt, 7)) Then
    ltxt = ltxt + " " + txt
    obAngeh = -1
   ElseIf obAngeh Then
    obAngeh = 0
   End If
  Else
   If Not rAdo!Lenge Like "###" Or Not rAdo!Kennung Like "####" Then
    RInhalt = RInhalt & " " & rAdo!Leng & rAdo!Kennung & rAdo!Inhalt
   ElseIf obAngeh Then
    obAngeh = 0
   End If
  End If
  
  Do
   If obAngeh Then
    If obEOF Then
     obSchluss = True
    Else
     Exit Do
    End If
   End If
    If LDInF = 0 Then
     RLng = Left(ltxt, 3)
     RKennung = Mid(ltxt, 4, 4)
     RInhalt = umw(Mid(ltxt, 8))
    Else
     If RInhalt <> "" Then
      RInhalt = umw(RInhalt)
     End If
    End If
    
    If ZS Then RInhalt = ZSU1(RInhalt)
    Call dolies(frm, RLng, RKennung, RInhalt, obSchluss) ' hierin auch alleSpeichern
    If obSchluss Or BrichAb Then Exit Do
   If LDInF = 0 Then
    ltxt = txt
   Else
    RInhalt = rAdo!Inhalt
    RKennung = rAdo!Kennung
    RLng = rAdo!Lenge
   End If
   If Not obEOF Then Exit Do
   obSchluss = True
  Loop
  If LDInF <> 0 Then rAdo.Move 1
  If BrichAb Then Exit Do
 Loop
 
 If obMySQL And obForeign Then Call DBCn.Execute("set foreign_key_checks= 1")

 Call EintragSpäter(2) ' Eintragsschleife verlassen
 If obSchluss Then
  Call EintragSpäter(3) ' ... und zwar an der vorgesehenen Stelle
 End If
' MedKlass, 2. Teil
 DBCn.BeginTrans
 If obMySQL And obForeign Then Call DBCn.Execute("set foreign_key_checks= 0")
 For i = 1 To MedSL.Count
  Dim Med$
  Med = MedSL.Item(i).Med
  If Med <> "" Then
   Set rs = DBCn.Execute("select * from medarten where medikament = '" & Med & "'")
   If rs.BOF Then
    Call DBCn.Execute("insert into medarten (medikament,hinzugefügt,langname,pat_id) values ('" & Med & "'," & datForm(Now) & ",'" & rMe(UBound(rMe)).Medikament & "'," & rNa(0).Pat_ID & ")")
   Else
    Set rs = DBCn.Execute("select * from medarten where medikament = '" & Med & "' and (pat_id = 0 or isnull(pat_id))")
    If Not rs.BOF Then
     Call DBCn.Execute("update medarten set pat_id = " & rNa(0).Pat_ID & " where medikament = '" & Med & "'", raf)
    End If
   End If
  End If
 Next i
 If obMySQL And obForeign Then Call DBCn.Execute("set foreign_key_checks= 1")
 DBCn.CommitTrans

#If False Then
 Call ProgrammLauf(-1)
#End If
'abbrech:
 Call EintragSpäter(4) ' nach MedKlass
 Call LaborDirektImport(frm)
 Call EintragSpäter(5)
 Call LaborErgPatId(frm)
 Call EintragSpäter(6)
 Call EmailsImport
 Call EintragSpäter(7)
 
 Call EintragSpäter(8)
 If LDInF = 0 Then
  Close #324
 End If
 Close #322
 frm.Ausgabe = "Zeitdauer gesamt: " & Format((Now - t1) * 60 * 60 * 24, "###,###,###,###,###,###,##0") & vbCrLf & frm.Ausgabe
 frm.Ausgabe = "Zeitdauer nach Vorladen: " & Format((Now - t1a) * 60 * 60 * 24, "###,###,###,###,###,###,##0") & vbCrLf & frm.Ausgabe
 For i = 0 To frm.Controls.Count - 1
  If frm.Controls(i).Name = "Beenden" Or frm.Controls(i).Name = "Zurücksetzen" Then
   frm.Controls(i).Enabled = True
  End If
'  If frm.Controls(i).Name = "Einlesen" Or frm.Controls(i).Name = "Abbrechen" Then
'   frm.Controls(i).Enabled = False
'  End If
 Next i
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in SpezKop/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GesLies
Function EintragStart(frm As Lese) ' wird in EinlesenClick aufgerufen
 On Error GoTo fehler
 Dim sql$
 rEzäBeg = Now
 Set rs = DBCn.Execute("select stbyte from eintragszahlen order by beginn desc")
 If rs.BOF Then
  VorByte = 0
 Else
  VorByte = rs.Fields(0)
 End If
 AktByte = VorByte + 1
 sql = "insert into eintragszahlen (beginn,stbyte,datei) values (" & datForm(rEzäBeg) & "," & AktByte & ",'" & umw(frm.QDatei) & "')"
 Call DBCn.Execute(sql)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in EintragStart/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' EintragStart() As DAO.Recordset
Function EintragSpäter(Nr%)
 On Error GoTo fehler
 Call DBCn.Execute("update eintragszahlen set zp" + CStr(Nr) + " = " + datForm(Now) + " where beginn = " + datForm(rEzäBeg) + "", numA)
 If numA <> 1 Then Err.Raise 999, , "Fehler in Eintragspäter: numA <> 1, und zwar: " & numA
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Eintrag/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' EintragSpäter

Function umw(Q$)
 If obMySQL Then
  umw = Replace(Replace(Trim(Q), "\", "\\"), "'", "\'")
 Else
  umw = Replace(Replace(Trim(Q), "\", "\\"), "'", "''")
 End If
End Function ' umw
Function StatischInit()
 Quartal = "" ' 4101
 lfdnr = 0
 lKennung = 0 ', lk%(15)
 jetztKopf = 0
 rFm_Nr = 0
 lFormID = 0  ' ID des letzten Formulars
 obZweiteRunde = 0
 lLang = ""
 lLangVW = 0
 lKomm = ""
 lKommVW = 0
 SpeicherZt = 0
 aktQ = ZQuart(Now - 11)
End Function ' StatischInit
Function dolies(frm As Lese, RLng$, RKennung$, RInhalt$, obSchluss%)
' If rAna Is Nothing Then
'  Set rAna = New ADODB.Recordset
'  rAna.Open "anamnesebogen", dbcn, adOpenDynamic, adLockOptimistic
' End If
 Dim FeldInhalt$, i&, j&
 Dim rFoNeu%, rFoAbkNeu% ' ob Formular/Abkürzungseintrag neu
 '  FID ggf zuweisen
 Dim obPat%
 Dim obNeuPat%
 On Error GoTo fehler
 absPos = absPos + 1
 
  AktZeit = Now ' 27.12.06: wird für die Tabelle Namen jetzt aus dem Feld 9103 gezogen
' Exit Function
 If RKennung Like "####" Then
'  Dim lPatNeu%
'  If lk(13) = 3000 Then Stop
'  If lKennung = 4101 Then Stop
  Select Case CLng(RKennung)
' Hier stand noch was
   Case 3000 ' Patientennummer
    If RInhalt <> rNa(0).Pat_ID Then
'     lPatNeu = -1
     If RInhalt = "0" Then
      If Not rsAdo Is Nothing Then If rsAdo.State = 1 Then rsAdo.Close
'      rsAdo.Open "select (max(pat_id) + 1) as pid from namen", dbcn, adOpenStatic, adLockReadOnly
      Set rsAdo = DBCn.Execute("select (max(pat_id) + 1) as pid from namen")
      Err.Raise 999, , "Achtung: Patientennummer 0 bei Datensatz " & RLng & "! Muß geändert werden, z.B. auf " + CStr(rsAdo!PID) + ": Bitte mit Turbomed abgleichen und auch dort ändern"
     End If
     obPat = True
     If rNa(0).Pat_ID <> 0 Then ' beim ersten Mal nicht
      obNeuPat = True
     End If
    Else
'     lPatNeu = 0
    End If
   Case 3100 ' Namensvorsatz
    rNa(0).NVorsatz = RInhalt
    rNa(0).AktZeit = LAktZeit ' Erstellung der BDT-Datei
 ' update macht dann den SQLBindParameter not used for all parameters mysq-fehler
 '  If rAna.EditMode = 1 Or rAna.EditMode = adEditInProgress Then
 '      rAna!NVorsatz = Rinhalt
 '   End If
   Case 3101 ' Nachname
    If RInhalt Like "zzz*" Then
     RInhalt = Mid(RInhalt, 4)
     Tkz = -1
    End If
    rNa(0).Nachname = RInhalt
   Case 3102 ' Vorname
    rNa(0).Vorname = RInhalt
   Case 3103 ' Geburtsdatum
    rNa(0).GebDat = BDTtoDate(RInhalt)
   Case 3104 ' Titel
    rNa(0).Titel = RInhalt
   Case 3105 ' Versichertennummer
    rNa(0).Versichertennummer = RInhalt
   Case 3107
    rNa(0).Straße = RInhalt
   Case 3108 ' KVKStatus
    rNa(0).KVKStatus = RInhalt
   Case 3110 ' Geschlecht
    rNa(0).Geschlecht = IIf(RInhalt = "1", "m", IIf(RInhalt = "2", "w", " "))
   Case 3112 ' Postleitzahl
    rNa(0).Plz = RInhalt
   Case 3113 ' Ort
    rNa(0).Ort = RInhalt
   Case 3610 ' Aufnahmedatum
    rNa(0).AufnDat = BDTtoDate(RInhalt)
    Call VorstellSetz(rNa(0).AufnDat)
   Case 3626 ' Telefonnummer von 3629 dupliziert
   Case 3629 ' verschiedene Telefonnummern mit Formatierung
'    Debug.Print RInhalt
    Call aufSplit(RInhalt)
    FeldInhalt = ""
    For i = 1 To ArraInd - 3
     FeldInhalt = FeldInhalt + Arra(i)
    Next
    Select Case Arra(ArraInd - 2)
     Case "1"
      Select Case Arra(ArraInd - 1)
       Case "Privat"
        If Arra(ArraInd) = "Tel" Then rNa(0).PrivatTel = FeldInhalt Else If Arra(ArraInd) = "Fax" Then rNa(0).PrivatFax = FeldInhalt Else If Arra(ArraInd) = "Mobil" Then rNa(0).PrivatMobil = FeldInhalt
       Case "Dienst"
        If Arra(ArraInd) = "Tel" Then rNa(0).DienstTel = FeldInhalt Else Err.Raise 999
       Case Else: Err.Raise 999
      End Select
     Case "2"
      Select Case Arra(ArraInd - 1)
       Case "Privat"
        If Arra(ArraInd) = "Tel" Then rNa(0).PrivatTel_2 = FeldInhalt Else Err.Raise 999
       Case Else: Err.Raise 999
      End Select
    End Select
   Case 3630 ' KV-Nummer des überweisenden
    If IsNull(RInhalt) Then
     Print #322, "Pat: " + CStr(rNa(0).Pat_ID) + ": KV-Nr des Überweisenden (Feld 3630) leer!"
    Else
     rNa(0).KVNr = Replace(RInhalt, "/", "")
     Dim obrKV%
     obrKV = -1
     For j = 1 To UBound(rKv)
      If rKv(j).KVNr = rNa(0).KVNr And rKv(j).Pat_ID = rNa(0).Pat_ID Then
       obrKV = 0
       Exit For
      End If
     Next j
     If obrKV Then
      ReDim Preserve rKv(UBound(rKv) + 1)
      rKv(UBound(rKv)).absPos = absPos
      rKv(UBound(rKv)).KVNr = rNa(0).KVNr
      rKv(UBound(rKv)).Pat_ID = rNa(0).Pat_ID
      rKv(UBound(rKv)).AktZeit = AktZeit
     End If
    End If
   Case 3631 ' Weggeldzone
    rNa(0).Weggeldzone = RInhalt
   Case 3632: rNa(UBound(rNa)).Verwandt = RInhalt
   Case 3633: rNa(UBound(rNa)).zubenach = RInhalt
   Case 3634: rNa(UBound(rNa)).Notiz = RInhalt
   Case 3635 ' interne Zuordnung Arzt bei GP: immer 1
    rNa(0).intZOGP = RInhalt
   Case 3649, 5999 ' Diagnosedatum
    messDatum = BDTtoDate(RInhalt)
   Case 3650, 6000 ' Dauerdiagnose
    ReDim Preserve rDi(UBound(rDi) + 1)
    rDi(UBound(rDi)).DiagDatum = messDatum
    rDi(UBound(rDi)).DiagText = RInhalt
    rDi(UBound(rDi)).FID = rFa(UBound(rFa)).FID
    rDi(UBound(rDi)).GesName = GesName(rNa(0).NVorsatz, rNa(0).Nachname, rNa(0).Titel, rNa(0).Vorname, rNa(0).GebDat)
    rDi(UBound(rDi)).absPos = absPos
    rDi(UBound(rDi)).AktZeit = AktZeit
    rDi(UBound(rDi)).Pat_ID = rNa(0).Pat_ID
    If RKennung = 3650 Then rDi(UBound(rDi)).obDauer = True
    
    If RKennung = 3650 Or rFa(UBound(rFa)).Quartal = aktQ Then ' formal müßte hier evtl altQuart rein, da dessen Belegung auf diese Schleife vorverlegt wurde
     diagDat(DiagNr) = messDatum
     Diag(DiagNr) = RInhalt
     For j = Len(Diag(DiagNr)) To 1 Step -1
      If Mid(Diag(DiagNr), j, 1) = "(" Then
       Dim IcdRoh$
'     If Pat_id = 33 Then Stop
       IcdRoh = Mid(Diag(DiagNr), j + 1, Len(Diag(DiagNr)) - j - 1)
       If IsNumeric(Mid(IcdRoh, 2, 1)) Then
        ICD(DiagNr) = IcdRoh
        Diag(DiagNr) = Left(Diag(DiagNr), j - 1)
       End If
       Exit For
      End If
     Next
     DiagSi(DiagNr) = "G"
     DiagSe(DiagNr) = ""
     DiagAttr(DiagNr) = ""
     obDauer(DiagNr) = rDi(UBound(rDi)).obDauer
     DiagNr = DiagNr + 1
     obDStr = True
    Else
     obDStr = False
    End If
   Case 3652, 6210 ' Medikament auf Rezept (neben 6210)
    Call RezEintr(RInhalt, False)
   Case 4101 ' Quartal
 '   If UBound(rfa) = 0 And rfa(0).Pat_ID = 0 Then ReDim Preserve rfa(UBound(rfa)) Else ReDim Preserve rfa(UBound(rfa) + 1)
    ReDim Preserve rFa(UBound(rFa) + 1)
    lfdfl = lfdfl + 1
    rFa(UBound(rFa)).s8000 = f8000
    f8000 = ""
    rFa(UBound(rFa)).s8100 = f8100
    f8100 = ""
    rFa(UBound(rFa)).Pat_ID = rNa(0).Pat_ID
    rFa(UBound(rFa)).Vorname = rNa(0).Vorname
    rFa(UBound(rFa)).Nachname = rNa(0).Nachname
    rFa(UBound(rFa)).lfdnr = lfdfl
    If Not rsAdo Is Nothing Then If rsAdo.State = 1 Then rsAdo.Close
    If UBound(rFa) = 1 Then
'     rsAdo.Open ("select max(fid) as fid from faelle")
     Set rsAdo = DBCn.Execute("select max(fid) as fid from faelle")
     If IsNull(rsAdo!FID) Then
      If obMySQL Then
       Call DBCn.Execute("alter table faelle auto_increment = 1")
      Else
       rsAdo.Close
       ZielDbS = frm.dlg.MdB
       Dim zwiCStr$
       zwiCStr = DBCn.ConnectionString
       DBCn.Close
       For i = 0 To 1000
       Next i
       Call BezLöschA
       For i = 0 To 1000
       Next i
       DBCn.Open zwiCStr
       Call DBCn.Execute("alter table faelle alter column fid counter(1,1)")
       Call BezHerstA
      End If
      rFa(UBound(rFa)).FID = 1
     Else
      rFa(UBound(rFa)).FID = nz(rsAdo!FID,0) + 1
     End If
    Else
     rFa(UBound(rFa)).FID = rFa(UBound(rFa) - 1).FID + 1
    End If
    rFa(UBound(rFa)).Quartal = RInhalt
    rFa(UBound(rFa)).QAnf = QAnf(Quartal)
    rFa(UBound(rFa)).QEnd = QEnd(Quartal)
    
   Case 4102 ' ausgestellt am
    rFa(UBound(rFa)).ausgst = BDTtoDate(RInhalt)
    Call VorstellSetz(BDTtoDate(RInhalt))
   Case 4104 ' VKNr
    rFa(UBound(rFa)).VKNr = RInhalt
   Case 4106 ' Kostenträgerabrechnungsbereich, immer 00
    rFa(UBound(rFa)).KtrAbrB = RInhalt
   Case 4107 ' Abrechnungsart, 1 oder 2
    rFa(UBound(rFa)).AbrAr = RInhalt
   Case 4109 ' letzte Vorlage
    rFa(UBound(rFa)).lVorl = BDTtoDate(RInhalt)
    Call VorstellSetz(rFa(UBound(rFa)).lVorl)
   Case 4110 ' Uhrzeit zu letzter Vorlage
    rFa(UBound(rFa)).lVorl = rFa(UBound(rFa)).lVorl + BDTtoTime(RInhalt)
   Case 4111 ' IK
    rFa(UBound(rFa)).IK = RInhalt
   Case 4112 'KVKStatus
    rFa(UBound(rFa)).KVKs = RInhalt
   Case 4113 ' Status-Ergänzung
    rFa(UBound(rFa)).KVKserg = RInhalt
   Case 4121 ' Gebührenordnung, 1 oder 2
    rFa(UBound(rFa)).GebOr = RInhalt
   Case 4122 ' Abrechnungsgebiet, 00
    rFa(UBound(rFa)).AbrGb = RInhalt
   Case 4144 ' Turbomed-Fallnummer
     Call aufSplit(RInhalt)
     If ArraInd > 0 Then
      rFa(UBound(rFa)).TMFNr = Arra(1)
     End If
   Case 4150 ' Behandlungsfall: Beginn
     rFa(UBound(rFa)).BhFB = BDTtoDate(RInhalt)
   Case 4151 ' Behandlungsfall: Ende (Musterwoman) / Wohl Ende des Behandlungsfallbeginnquartals
     rFa(UBound(rFa)).BhFE1 = BDTtoDate(RInhalt)
   Case 4152 ' Behandlungsfall: Ende (Musterwoman)
     rFa(UBound(rFa)).BhFE2 = BDTtoDate(RInhalt)
   Case 4202 ' Unfall, Unfallfolgen nach 4152, bei Maurer, Temiz und Schilcher "1"
     rFa(UBound(rFa)).f4202 = RInhalt
   Case 4206 ' mutmaßlicher Tag der Entbindung
     rFa(UBound(rFa)).f4206 = RInhalt
   Case 4209 ' Überweisungstext
     rFa(UBound(rFa)).ÜwText = RInhalt
   Case 4210 ' unbekannt
     rFa(UBound(rFa)).f4210 = RInhalt
   Case 4218 ' Überwiesen von
     rFa(UBound(rFa)).ÜbwV = IIf(IsNull(RInhalt), "", Replace(RInhalt, "/", ""))
   Case 4219 ' Anderer Überweiser
     rFa(UBound(rFa)).AndÜw = IIf(IsNull(RInhalt), "", Replace(RInhalt, "/", ""))
   Case 4220 ' Überweisungsziel
     rFa(UBound(rFa)).ÜWZiel = RInhalt
   Case 4231, 4241 ' Überweiser, hat sich offenbar von 4231 auf 4241 geändert! (31.7.05)
     Call aufSplit(RInhalt)
     rFa(UBound(rFa)).ÜWNNr = IIf(IsNull(Arra(ArraInd)), "", Replace(Arra(ArraInd), "/", ""))
     rFa(UBound(rFa)).ÜWNaN = Arra(ArraInd - 1)
     rFa(UBound(rFa)).ÜWVsw = Arra(ArraInd - 2)
     rFa(UBound(rFa)).ÜWVor = Arra(ArraInd - 3)
     rFa(UBound(rFa)).ÜWTit = Arra(ArraInd - 4)
   Case 4233
     rFa(UBound(rFa)).statNuller = RInhalt
   Case 4236 ' bei stationär: Klasse
     rFa(UBound(rFa)).statKlasse = RInhalt
   Case 4237 ' unbek
     rFa(UBound(rFa)).f4237 = RInhalt
   Case 4238 ' bei stationär: Behandlungstage
     rFa(UBound(rFa)).statBehTage = RInhalt
   Case 4239 ' Scheingruppe, 24, 41, 90
     rFa(UBound(rFa)).SchGr = RInhalt
     ' zuvor: testaltQuartal
     Select Case rFa(UBound(rFa)).SchGr
      Case "00", "20", "21", "23", "24"  ' 23 = kurativ Konsiliaruntersuchung 21 = Auftragsleistung
       If rFa(UBound(rFa)).Quartal <> ZQuart(rFa(UBound(rFa)).BhFB) Then
        rFa(UBound(rFa)).altQuart = rFa(UBound(rFa)).Quartal
        rFa(UBound(rFa)).Quartal = ZQuart(rFa(UBound(rFa)).BhFB)
       End If
      Case "90", "92", "41", "43" ' 43 = Notfall 41 = ÄND, 92 = nur Musterwoman
      Case Else
'    Debug.Print aFäl!Pat_id, aFäl!Nachname, aFäl!Vorname, aFäl!SchGr, aFäl!FID, aFäl!Quartal, aFäl!BhFB
    End Select
   Case 4243 ' Weiterbehandelnder Arzt in Worten
     rFa(UBound(rFa)).Weiterbeh = RInhalt
   Case 4401 ' 4401, Praxisgebühr -Array
     Call aufSplit(RInhalt)
     rFa(UBound(rFa)).PGeb = Arra(1)
   Case 4402 ' PGebErg  '4402, Array ?
     Call aufSplit(RInhalt)
     rFa(UBound(rFa)).PGebErg = Arra(1)
   Case 4403 ' 4403, Mahnfrist bis
     Call aufSplit(RInhalt)
     rFa(UBound(rFa)).Mahnfrist = Arra(1)
   Case 4580 ' Abrechnungsart
     Call aufSplit(RInhalt)
     If ArraInd > 0 Then rFa(UBound(rFa)).GOÄKatNr = Arra(1)
     If ArraInd > 1 Then rFa(UBound(rFa)).GOÄKatName = Arra(2)
   Case 4602 ' Adresse bei Musterfrau
     Call aufSplit(RInhalt)
     If ArraInd > -1 Then rFa(UBound(rFa)).AdNam = Arra(0)
     If ArraInd > 0 Then rFa(UBound(rFa)).AdStr = Arra(1)
     If ArraInd > 1 Then rFa(UBound(rFa)).AdPlz = Arra(2)
     If ArraInd > 2 Then rFa(UBound(rFa)).AdOrt = Arra(3)
   Case 4604  'Behandlungsfall: Ende (Musterwoman)
     rFa(UBound(rFa)).BhFE = BDTtoDate(RInhalt)
   Case 5000, 6200 ' Datum
    messDatum = BDTtoDate(RInhalt)
' Hier fängt auf jeden Fall ein neues Formular an
    FormID = 0
    FormAbk = ""
    DokPfad = ""
    DokArt = ""
    DokName = ""
   Case 5001 ' Leistungsziffer
    Call LeistEintr0(RInhalt)
    Select Case UCase(RInhalt)
     Case "9261B", "9262B", "9263B", "9264B", "9268", "9269", "9270", "9271", _
          "97261B", "97262B", "97263B", "97264B", "97268", "97269", "97270", "97271"
      DMSchulz = DMSchulz + 1
     Case "9265B", "9266B", _
          "97265B", "97266B", "97268B"
      RRSchulz = RRSchulz + 1
    End Select
    If Not obvorgestellt Then
     Call VorstellSetz(messDatum)
    End If
   Case 5002 ' Art der Untersuchung
    rLe(UBound(rLe)).f5002 = RInhalt
   Case 5005 ' Anzahl
    rLe(UBound(rLe)).f5005 = RInhalt
   Case 5006 ' um Uhrzeit
    rLe(UBound(rLe)).f5006 = RInhalt
   Case 5009 ' freier Text
    rLe(UBound(rLe)).f5009 = RInhalt
   Case 5010 ' Medikament
    rLe(UBound(rLe)).Med = RInhalt
   Case 5015 ' Organ
    rLe(UBound(rLe)).f5015 = RInhalt
   Case 5016 ' Name des Arztes (Briefempfänger)
    rLe(UBound(rLe)).f5016 = RInhalt
   Case 5021 ' Datum letzte Krebsvorsorge
    rLe(UBound(rLe)).f5021 = RInhalt
   Case 5026 ' Entlassungsdatum
    rLe(UBound(rLe)).f5026 = RInhalt
   Case 6001 ' ICD
    rDi(UBound(rDi)).ICD = RInhalt
    If obDStr Then ICD(DiagNr - 1) = RInhalt
   Case 6003  ' Diagnosensicherheit
    rDi(UBound(rDi)).DiagSicherheit = RInhalt
    If obDStr Then DiagSi(DiagNr - 1) = RInhalt
   Case 6004 ' Diagnosenseite
    rDi(UBound(rDi)).DiagSeite = RInhalt
    If obDStr Then DiagSe(DiagNr - 1) = RInhalt
   Case 6006 ' Diagnosenattribut
    rDi(UBound(rDi)).DiagAttr = RInhalt
    If obDStr Then DiagAttr(DiagNr - 1) = RInhalt
   Case 6201 ' Uhrzeit
    messDatum = messDatum + BDTtoTime(RInhalt)
   Case 6218 ' Rezepteintrag
    Call RezEintr(RInhalt, True)
   Case 6220 ' Symptome, bisher nur bei Musterfrau
   Case 6230 ' Blutdruck
    Call RREintr(RInhalt)
    DStr = " (" + Format(messDatum, "dd/mm/yy") + ")"
    If Arr = "" Then Arr = RInhalt + DStr
    lRR = RInhalt + DStr
    gesRR = gesRR + IIf(gesRR = "", "", ", ") + RInhalt + DStr
   Case 6280 ' Laboranforderungstext
    If lKennung = 6280 Then
     rLb(UBound(rLb)).AnfText = Left(rLb(UBound(rLb)).AnfText, Len(rLb(UBound(rLb)).AnfText) - IIf(Right(rLb(UBound(rLb)).AnfText, 1) = "^", 1, 0)) + RInhalt
    Else
     ReDim Preserve rLb(UBound(rLb) + 1)
     rLb(UBound(rLb)).Pat_ID = rNa(0).Pat_ID
     rLb(UBound(rLb)).Zeitpunkt = messDatum
     rLb(UBound(rLb)).AnfText = RInhalt
     rLb(UBound(rLb)).FID = rFa(UBound(rFa)).FID
     rLb(UBound(rLb)).absPos = absPos
     rLb(UBound(rLb)).AktZeit = AktZeit
    End If
   Case 6285 ' AU-Dauer
    ReDim Preserve rAu(UBound(rAu) + 1)
    rAu(UBound(rAu)).Pat_ID = rNa(0).Pat_ID
    rAu(UBound(rAu)).Zeitpunkt = messDatum
    rAu(UBound(rAu)).Beginn = Left(RInhalt, 8)
    rAu(UBound(rAu)).Ende = Right(RInhalt, 8)
    rAu(UBound(rAu)).FID = rFa(UBound(rFa)).FID
    rAu(UBound(rAu)).AktZeit = AktZeit
   Case 6286 ' AU-Begründung
    If rAu(UBound(rAu)).ICDs = "" Then
     rAu(UBound(rAu)).ICDs = RInhalt
    Else
     rAu(UBound(rAu)).ICDs = rAu(UBound(rAu)).ICDs + " " + RInhalt
    End If
   Case 6291 ' Krankenhauseinweisung
    ReDim Preserve rKh(UBound(rKh) + 1)
    rKh(UBound(rKh)).Pat_ID = rNa(0).Pat_ID
    rKh(UBound(rKh)).Zeitpunkt = messDatum
    rKh(UBound(rKh)).Ziel = RInhalt
    rKh(UBound(rKh)).absPos = absPos
    rKh(UBound(rKh)).FID = rFa(UBound(rFa)).FID
    rKh(UBound(rKh)).AktZeit = AktZeit
   Case 6295 ' Formularabkürzung
    FormAbk = RInhalt
   Case 6296 ' Formular-Bezeichnung
    FormBez = RInhalt
    If RInhalt = "Medikamentenplan" Then
     If obMedPlanGelesen Then
      obMedPlan = False
     Else
      obMedPlan = True
      obMedPlanGelesen = True
     End If
    End If
   Case 6297 ' Formular-Vorlage (Windows-Pfad)
    FormVorl = RInhalt 'Replace(RInhalt, "€", " ")
    If FormBez = "Medikamentenplan" Then
     MedZahl = -1
     If MPNr = 0 Then
      If Not rsAdo Is Nothing Then If rsAdo.State = 1 Then rsAdo.Close
'      rsAdo.Open "select (max(mpnr) + 1) as mpnr from medplan", dbcn, adOpenStatic, adLockReadOnly
      Set rsAdo = DBCn.Execute("select (max(mpnr) + 1) as mpnr from medplan")
      MPNr = nz(rsAdo!MPNr,1)
     Else
      MPNr = MPNr + 1
     End If
    Else ' FormBez = "Medikamentenplan"
     rFoNeu = -1
     For i = 1 To UBound(rFo)
      If rFo(i).Form_Abk = FormAbk And rFo(i).FormBez = FormBez Then
       If umw(rFo(i).FormVorl) = FormVorl Or rFo(i).FormVorl = FormVorl Then
        rFoNeu = 0
        lFormID = rFo(i).FormID
        Exit For
       Else
'        stop
       End If
      End If
     Next i
     If rFoNeu Then
      ReDim Preserve rFo(UBound(rFo) + 1)
      rFo(UBound(rFo)).Form_Abk = FormAbk
      rFo(UBound(rFo)).absPos = absPos
      rFo(UBound(rFo)).AktZeit = AktZeit
      rFo(UBound(rFo)).FormBez = FormBez
      rFo(UBound(rFo)).FormID = rFo(UBound(rFo) - 1).FormID + 1 ' muss noch in FormInhKopf angpasst werden, wird deshalb dort negativ gespeichert
      lFormID = rFo(UBound(rFo)).FormID
      rFo(UBound(rFo)).FormVorl = FormVorl
      rFo(UBound(rFo)).StByte = AktByte
     End If
     rFoAbkNeu = -1
     For i = 1 To UBound(rFi)
      If UCase(rFi(i).Form_Abk) = UCase(FormAbk) Then
       rFoAbkNeu = 0
       Exit For
      End If
     Next i
     If rFoAbkNeu Then
      ReDim Preserve rFi(UBound(rFi) + 1)
      rFi(UBound(rFi)).Form_Abk = FormAbk
     End If
     rFm_Nr = 0
     jetztKopf = True
    End If ' FormBez = "Medikamentenplan"
   Case 6298 ' Formularfeld und ggf. Spalte
    Call aufSplit(RInhalt)
    If FormBez = "Medikamentenplan" Then
'     MedArt = Arra(0)
     MedNr = -1
     MedZeit = -1
'     ReDim Med$(0)
'     ReDim Dos$(5, 0)
     Select Case Arra(0)
      Case "medikament"
       MedNr = Arra(1)
       While CLng(Arra(1)) > MedZahl
        Call MPerg
       Wend
      Case Else
'       Debug.Print Arra(0)
'       Stop
     End Select ' arra(0)
    Else
     If jetztKopf Then
      ReDim Preserve rFr(UBound(rFr) + 1)
      rFr(UBound(rFr)).absPos = absPos
      rFr(UBound(rFr)).AktZeit = AktZeit
      rFr(UBound(rFr)).FID = rFa(UBound(rFa)).FID
      rFr(UBound(rFr)).Form_ID = lFormID '-lFormID ' negative Speicherung, da der Wert noch nach der Datenbankspeicherung von rFo angepaßt werden muss
      rFr(UBound(rFr)).Pat_ID = rNa(0).Pat_ID
      rFr(UBound(rFr)).StByte = AktByte
      rFr(UBound(rFr)).Zeitpunkt = messDatum
      rFr(UBound(rFr)).Satzart = f8000
      f8000 = ""
      rFr(UBound(rFr)).Satzlänge = f8100
      f8100 = ""
      If FoIDv = 0 Then
       If Not rsAdo Is Nothing Then If rsAdo.State = 1 Then rsAdo.Close
       Set rsAdo = DBCn.Execute("select (max(foid) + 1) as mfoid from forminhkopf")
       FoIDv = nz(rsAdo!mfoid,0)
      End If
      rFr(UBound(rFr)).FoID = FoIDv ' Pseudo-Foid
      FoIDv = FoIDv + 1
      rFm_Nr = rFm_Nr + 1
      jetztKopf = False
     End If
    End If ' Formbez
    ZeileNr = 0
   Case 6299 ' Formularfeldinhalt auch Überweisungstext
    FormInh = RInhalt
    If FormBez = "Medikamentenplan" Then
    Else
     ReDim Preserve rFm(UBound(rFm) + 1)
     rFm(UBound(rFm)).FeldInh = FormInh
     rFm(UBound(rFm)).Feld = Arra(0)
     If ArraInd > 0 Then
      rFm(UBound(rFm)).FeldNr = Arra(1)
     Else
      rFm(UBound(rFm)).FeldNr = ZeileNr ' damit mehrere 6299 auf ein 6298 kommen können
     End If
     rFm(UBound(rFm)).Nr = rFm_Nr
     rFm_Nr = rFm_Nr + 1
     rFm(UBound(rFm)).FoID = FoIDv - 1 '-FoID 'negative Speicherung, da der Wert noch nach der Datenbankspeicherung von rFr angepaßt werden muss
    End If
'    Call FormInhEintrag(rtr.AbsolutePosition)
    If FormBez Like "Erst-Dokumentation Diabetes mellitus Typ 2*" Then
    ElseIf FormBez Like "Folge-Dokumentation Diabetes mellitus Typ 2*" Then
    Else
     Select Case FormBez
      Case "Medikamentenplan"
       If Arra(0) = "Datum" Then
        MPDatum = FormInh
        For i = 0 To MedZahl
         rMe(UBound(rMe) - i).Datum = MPDatum
        Next i
       ElseIf MedNr <> -1 Then
        rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).Medikament = FormInh
        rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).MedAnfang = Left(FormInh, IIf(InStr(FormInh, " ") > 0, InStr(FormInh, " "), Len(FormInh)))
       ElseIf MedZahl <> -1 Or Arra(0) = "Check1" Then
        If ArraInd > 0 Then
         While CLng(Arra(1)) > MedZahl
          Call MPerg
         Wend
        End If
        Select Case Arra(0)
         Case "Check1"
          If FormInh = "X" Then rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).bBed = True
         Case "Hergang"
          rMe(UBound(rMe) - MedZahl).Bemerkung = LTrim(rMe(UBound(rMe) - MedZahl).Bemerkung + " ") + FormInh
         Case "mo"
          rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).mo = FormInh
         Case "mi"
          rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).mi = FormInh
         Case "nm"
          rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).nm = FormInh
         Case "ab"
          rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).ab = FormInh
         Case "zn"
          rMe(UBound(rMe) - MedZahl + CLng(Arra(1))).zn = FormInh
         Case "Datum"
          MPDatum = FormInh
         Case "Nachname", "Vorname", "Name", "Stempel", "Datenfeld1602", "Datenfeld40€1€0", "Datenfeld40_1_0", "geb", "geb€0", "geb_0", "geb€0€0", "geb_0_0", "geb€0€0€0", "geb_0_0_0", "geb€0€0€0€0", "geb_0_0_0_0", "geb€0€0€0€0€0", "geb_0_0_0_0_0", "geb€0€0€1", "geb_0_0_1", "geb€0€1", "geb_0_1"
         Case Else
          Err.Raise 999, , "Fehler in dolies: Arra(0) im Medplan hat den unvorhergesehenen Wert: " & Arra(0)
        End Select
       End If
      Case "Privatrezept"
      Case "Rezept"
      Case "Überweisung Labor"
      Case "GDT"
      Case "Lang-Rezept"
      Case "Sprechstundenbedarf Medikamente"
      Case "Sprechstundenbedarf BTM"
      Case "Sprechstundenbedarf Heil- und Hilfsmittel"
      Case "Überweisung gezielt an: Augenheilkunde"
      Case "Überweisung gezielt an: Psychologie"
      Case "Überweisung gezielt an: Dermatologie"
      Case "Überweisung gezielt an: Psychotherapie"
      Case "Überweisung gezielt an: Gynäkologie"
      Case "Überweisung gezielt an: ->Psychotherapie"
      Case "Überweisung gezielt an: Neurologie"
      Case "Heilmittelverordnung Physikalische Therapie"
      Case "Überweisung gezielt an: Orthopädie"
      
'     Stop
      Case Else
       Print #322, "FormBez: " + FormBez + " in do_DatenEintragen bisher nicht berücksichtigt!"
       If InStr(FormBez, "EKG -GDT-Meßwert") > 0 Then
'      Stop
       End If
'     Debug.Print "Formbezeichnung: " + FormBez
     End Select
    End If
    If FormAbk = "DMPDTYP2" Then If DMPDat = 0 Or DMPDat < messDatum Then DMPDat = messDatum
    ZeileNr = ZeileNr + 1
   Case 6324 ' Brief mit Pfad
    Call aufSplit(RInhalt)
    ReDim Preserve rBr(UBound(rBr) + 1)
    rBr(UBound(rBr)).absPos = absPos
    rBr(UBound(rBr)).AktZeit = AktZeit
    If ArraInd > -1 Then rBr(UBound(rBr)).Pfad = Arra(0)
    If ArraInd > 0 Then rBr(UBound(rBr)).Art = Arra(1)
    If ArraInd > 1 Then rBr(UBound(rBr)).Name = Arra(2)
    If ArraInd > 2 Then rBr(UBound(rBr)).Typ = Arra(3)
    rBr(UBound(rBr)).FID = rFa(UBound(rFa)).FID
    rBr(UBound(rBr)).Pat_ID = rNa(0).Pat_ID
    rBr(UBound(rBr)).Zeitpunkt = messDatum
    rBr(UBound(rBr)).QT = ZQuart(rBr(UBound(rBr)).Zeitpunkt)
    rBr(UBound(rBr)).QS = ZQuartSort(rBr(UBound(rBr)).Zeitpunkt)
    If PcDokPfad = "" Then getDokPfad
    
    Dim Datei$
    Datei = Replace(rBr(UBound(rBr)).Pfad, "$\\TurboMed\\Dokumente", PcDokPfad)
    If FSO.FileExists(Datei) Then
     Set Fil = FSO.GetFile(Datei)
     If obMySQL Then
      rBr(UBound(rBr)).Pfad = Replace(Fil.Path, "\", "\\")
     Else
      rBr(UBound(rBr)).Pfad = Fil.Path
     End If
     If Not Fil Is Nothing Then
      rBr(UBound(rBr)).DokGroe = Fil.Size
     End If
    End If
   Case 6325 ' Dokumentpfad
    If Not obZweiteRunde Then
     ReDim Preserve rDo(UBound(rDo) + 1)
    End If
    rDo(UBound(rDo)).DokPfad = rDo(UBound(rDo)).DokPfad + Stutz(RInhalt)
    If Right(RInhalt, 1) = "^" Then
     obZweiteRunde = -1
    Else
     obZweiteRunde = 0
     Datei = Replace(rDo(UBound(rDo)).DokPfad, "$\\TurboMed\\Dokumente", PcDokPfad)
     If FSO.FileExists(Datei) Then
      Set Fil = FSO.GetFile(Datei)
      If obMySQL Then
       rDo(UBound(rDo)).DokPfad = Replace(Fil.Path, "\", "\\")
      Else
       rDo(UBound(rDo)).DokPfad = Fil.Path
      End If
      If Not Fil Is Nothing Then
       rDo(UBound(rDo)).DokGroe = Fil.Size
      End If
     End If
    End If
   Case 6326 ' Dokumentart (Windows-Dateityp)
    rDo(UBound(rDo)).DokArt = RInhalt
   Case 6327, 6300 ' Dokumentname / 6300 = ohne Ortsangabe (11.12.06: 6300 kommt nicht (mehr) vor)
    rDo(UBound(rDo)).DokName = rDo(UBound(rDo)).DokName + Stutz(RInhalt)
    rDo(UBound(rDo)).absPos = absPos
    rDo(UBound(rDo)).AktZeit = AktZeit
    rDo(UBound(rDo)).FID = rFa(UBound(rFa)).FID
    rDo(UBound(rDo)).Pat_ID = rNa(0).Pat_ID
    rDo(UBound(rDo)).Zeitpunkt = messDatum
    rDo(UBound(rDo)).QS = ZQuartSort(rDo(UBound(rDo)).Zeitpunkt)
    rDo(UBound(rDo)).QT = ZQuart(rDo(UBound(rDo)).Zeitpunkt)
    rDo(UBound(rDo)).Quelldatum = doQuelldatum(rDo(UBound(rDo)).DokName)
   Case 6330 ' Eintraege: Art
    If RInhalt = "DiagTxt" Then
     obEinweisung = True
    Else
     obEinweisung = False
     ReDim Preserve rEi(UBound(rEi) + 1)
     rEi(UBound(rEi)).Art = RInhalt
     rEi(UBound(rEi)).absPos = absPos
     rEi(UBound(rEi)).AktZeit = AktZeit
     rEi(UBound(rEi)).FID = rFa(UBound(rFa)).FID
     rEi(UBound(rEi)).Pat_ID = rNa(0).Pat_ID
     rEi(UBound(rEi)).QS = ZQuartSort(rDo(UBound(rDo)).Zeitpunkt)
     rEi(UBound(rEi)).QT = ZQuart(rDo(UBound(rDo)).Zeitpunkt)
     rEi(UBound(rEi)).Zeitpunkt = messDatum
     If RInhalt Like "*. sitzung*" Then DMSchL = DMSchL + 1
    End If
   Case 6331 ' Eintraege: Inhalt
'    EintrInh = EintrInh + Stutz(RInhalt)
    If obEinweisung Then
     rKh(UBound(rKh)).Diagnose = rKh(UBound(rKh)).Diagnose + Stutz(RInhalt)
    Else
     rEi(UBound(rEi)).Inhalt = rEi(UBound(rEi)).Inhalt + Stutz(RInhalt)
     Call PraxisHbA1c(rEi(UBound(rEi)).Art, rEi(UBound(rEi)).Inhalt)
     If rEi(UBound(rEi)).Art = "schul" And InStr(rEi(UBound(rEi)).Inhalt, "schulungsverein") > 0 Then
      obMedNetz = True
     End If
    End If
    If RInhalt = "bfd" And lKennung = 6330 And InStr(RInhalt, "icral") > 0 Then
     AlbErhöht = RInhalt
     AlbDat = messDatum
    End If
  Case 8000 ' Satzart (Turbomed) / Satzstatus (Internet-PDF-Dokument): 0102 macht Kassenfall, 0190 Privatfall
    f8000 = RInhalt
  Case 8100 ' Satzlänge (?)
    f8100 = RInhalt
  Case 8401 ' Labor: Fertigstellungsgrad
    FStG = RInhalt
'    ReDim Preserve rLa(UBound(rLa) + 1)
'    rLa(UBound(rLa)).FertigStGrad = RInhalt
  Case 8410 ' Labor: Abkürzung
    Abkü = RInhalt 'Replace(RInhalt, "€", "_") ' ZS1 wäre hier nötig
'    rLa(UBound(rLa)).Abkü = RInhalt
  Case 8411 ' Labor: Langtext
    Langtext = Replace(RInhalt, "³", "ü") ' scheint die einzig angezeigte Ersetzung aus Zeichensatz() zu sein
'    rLa(UBound(rLa)).Langtext = ZeichenSatz(RInhalt)
  Case 8420 ' Labor: Wert
    Wert = RInhalt
'    rLa(UBound(rLa)).Wert = RInhalt
'    If rLa(UBound(rLa)).Abkü = "HBA1C" Then
    If Abkü = "HBA1C" Then
     If messDatum > lHbA1cDat Then
      lHbA1c = RInhalt
      lHbA1cDat = messDatum
     End If
    End If
  Case 8421
   Einheit = Replace(RInhalt, "³", "ü")  ' scheint die einzig angezeigte Ersetzung aus Zeichensatz() zu sein
'   rLa(UBound(rLa)).Einheit = RInhalt
    Call LaborEintr0
  Case 8480
    Dim p1%, p2%
'   rLa(UBound(rLa)).Kommentar = rLa(UBound(rLa)).Kommentar + Stutz(RInhalt)
   Kommentar = Kommentar + Stutz(RInhalt)
' jetzt kommen noch die nicht-numerischen Werte
    If Wert = "" Then
     If InStr(RInhalt, "<") > 0 Then
      Wert = Mid(RInhalt, InStr(RInhalt, "<"))
      If InStr(4, Wert, " ") > 0 Then
       Wert = Left(Wert, InStr(4, Wert, " ") - 1)
      End If
     ElseIf InStr(RInhalt, ">") > 0 Then
      Wert = Mid(RInhalt, InStr(RInhalt, ">"))
      If InStr(4, Wert, " ") > 0 Then
       Wert = Left(Wert, InStr(4, Wert, " ") - 1)
      End If
     ElseIf InStr(RInhalt, ":") > 0 Then ' 1:-Größen ohne < und >
      p1 = InStr(RInhalt, ":")
      For p2 = p1 To 1 Step -1
       If Mid(RInhalt, p2, 1) = " " Then Exit For
      Next p2
      Wert = Mid(RInhalt, p2 + 1)
      If InStr(Wert, " ") > 0 Then
       Wert = Left(Wert, InStr(Wert, " ") - 1) ' Erklärungen weg
      End If
     End If
    End If ' wert=""
    Call Laboreintr1
  Case 4585 ' wahrscheinlich der abrechnende Arzt
'    Stop
  Case 9100 ' Arztnummer des Absenders
  Case 9103 ' Erstellungsdatum
'     NAktZeit = BDTtoDate(RInhalt)
  Case 9105 ' Ordnungsnummer Datenträger (Header) des DP
  Case 9106 ' verwendeter Zeichensatz
  Case 9210 ' Version ADT-Satzbeschreibung
  Case 9213 ' Version BDT-Satzbeschreibung
  Case 9600 ' Archivierungsart
  Case 9601 ' Zeitraum der Speicherung
   SpeicherZt = BDTtoDate(Right(RInhalt, 8))
  Case 9602 ' Beginn der Übertragung
   SpeicherZt = SpeicherZt + BDTtoTime(Left(RInhalt, 6))
   Call DBCn.Execute("update eintragszahlen set speicherzt = " + datForm(SpeicherZt) + " where beginn = " + datForm(rEzäBeg) + "", numA)
  Case 4211 ' ?
  Case 5010 ' Medikament
  Case 9202 ' Gesamtlänge des Datenpaketes
  Case 9203 ' Anzahl der Datenträger des Datenpaketes
  Case 101 ' KBV -Prüfnummer            : A001011
  Case 102 ' Software -Verantwortlicher : Turbomed EDV Gmbh
  Case 103 ' Software                   : Turbomed@Windows
  Case 104 '                            : IBM PC/AT
  Case 201 ' Arztnummer                 : 6419153
  Case 202 ' Praxistyp                  : 1 (Einzelpraxis)
  Case 203 ' Arztname                   : Gerald Schade
  Case 204 ' Arztgruppe                 : Innere Medizin
  Case 205 ' Straße der Praxis          : Münchner Straße 61b
  Case 215 ' PLZ der Praxisadresse      : 85221
  Case 216 ' Ort der Praxisadresse      : Dachau
  Case 208 ' Telefonnummer der Praxis   : 08131 / 616 380
  Case 209 ' Telefaxnummer der Praxis   : 08131 / 616 381
  Case 101, 102, 103, 104 ' Installation
  Case 201, 202, 203, 204, 205, 215, 216, 208, 209 ' KV-Nummer, meine Adresse
  Case Else ' Unbekannte Kennung
    ReDim Preserve rUn(UBound(rUn) + 1)
    rUn(UBound(rUn)).absPos = absPos
    rUn(UBound(rUn)).Kennung = RKennung
 End Select
 lKennung = RKennung
'  For i = UBound(lk) To 1 Step -1
'   lk(i) = lk(i - 1)
'   lk(0) = Rkennung
'  Next i
 End If ' Rkennung Like "####" Then
 If obNeuPat Or obSchluss Then
  frm.Ausgabe = lfdnr & ") " & Format(VorStDat, "dd.mm.yy") & ": (" & rNa(0).Pat_ID & ") " & rNa(0).Nachname & " " & rNa(0).Vorname & ", " & rNa(0).GebDat & ";" & vbCrLf + frm.Ausgabe
  Call alleSpeichern(frm)
  Call Tinit
 End If
 If obPat Then
  lfdnr = lfdnr + 1
  rNa(0).lfdnr = lfdnr
  rNa(0).Pat_ID = RInhalt
  Call PatInit
 End If
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in dolies/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' dolies
Function MPerg()
       ReDim Preserve rMe(UBound(rMe) + 1)
       MedZahl = MedZahl + 1
       rMe(UBound(rMe)).FeldNr = MedZahl
       rMe(UBound(rMe)).absPos = absPos
       rMe(UBound(rMe)).StByte = AktByte
       rMe(UBound(rMe)).AktZeit = AktZeit
       rMe(UBound(rMe)).Pat_ID = rNa(0).Pat_ID
       rMe(UBound(rMe)).FID = rFa(UBound(rFa)).FID
       rMe(UBound(rMe)).Zeitpunkt = messDatum
       rMe(UBound(rMe)).Datum = MPDatum
       rMe(UBound(rMe)).MPNr = MPNr
End Function ' MPerg
Function QAnf(Quartal) As Date ' Wird zZt nur für falleintragen gebraucht, nicht in Abfragen, da die zugehörige Abfrage zu langsam ist
Dim Q As String * 1, DatStr$
If Quartal = "" Then
 QAnf = CDate(0)
Else
 Q = Left(Quartal, 1)
 Select Case Q
  Case 1: DatStr = "1.1."
  Case 2: DatStr = "1.4."
  Case 3: DatStr = "1.7."
  Case 4: DatStr = "1.10."
 End Select
 QAnf = CDate(DatStr + Right(Quartal, 4))
End If
End Function ' QAnf(quartal) As Date ' Wird zZt nicht gebraucht, da die zugehörige Abfrage zu langsam ist
Function QEnd(Quartal) As Date ' Wird zZt nur für falleintragen gebraucht, nicht in Abfragen, da die zugehörige Abfrage zu langsam ist
Dim Q As String * 1, DatStr$
If Quartal = "" Then
 QEnd = CDate(0)
Else
 Q = Left(Quartal, 1)
 Select Case Q
  Case 1: DatStr = "31.3."
  Case 2: DatStr = "30.6."
  Case 3: DatStr = "30.9."
  Case 4: DatStr = "31.12."
 End Select
 QEnd = CDate(DatStr + Right(Quartal, 4))
End If
End Function ' QEnd(quartal) As Date ' Wird zZt nicht gebraucht, da die zugehörige Abfrage zu langsam ist
Function LeistEintr0(lg)
 On Error GoTo fehler
 If Not IsNull(lg) Then
  ReDim Preserve rLe(UBound(rLe) + 1)
  rLe(UBound(rLe)).Pat_ID = rNa(0).Pat_ID
  rLe(UBound(rLe)).Zeitpunkt = messDatum
  rLe(UBound(rLe)).Leistung = lg
  rLe(UBound(rLe)).QS = ZQuartSort(messDatum)
  rLe(UBound(rLe)).QT = ZQuart(messDatum)
  rLe(UBound(rLe)).absPos = absPos
  rLe(UBound(rLe)).AktZeit = AktZeit
  rLe(UBound(rLe)).FID = rFa(UBound(rFa)).FID
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LeistEintr0/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LeistEintr(Lg)
Function test7()
 Dim rs As New ADODB.Recordset, dt$
 Set rs = DBCn.Execute("select * from eintraege where pat_id = 853 and art = ""andm""")
 Debug.Print rs!Inhalt
 dt = Mid(rs!Inhalt, InStr(rs!Inhalt, "Diabetes Typ") + 13, InStr(rs!Inhalt, "seit") - InStr(rs!Inhalt, "Diabetes Typ") - 13)
 Debug.Print dt
End Function ' test7
Function fDTyp$()
' DTypeintrag
 Dim dt$, i& ', erg$ ', rAn As New ADODB.Recordset
 For i = 1 To UBound(rEi)
  If rEi(i).Art = "andm" Then
   dt = Mid(rEi(i).Inhalt, InStr(rEi(i).Inhalt, "Diabetes Typ") + 13, InStr(rEi(i).Inhalt, "seit") - InStr(rEi(i).Inhalt, "Diabetes Typ") - 13)
   If InStr(dt, "{") > 0 Then
    dt = Left(dt, InStr(dt, "{") - 1)
   End If
   dt = Trim(dt)
'   If Not rAn Is Nothing Then If rAn.State = 1 Then rAn.Close
'   rAn.Open "select * from anamnesebogen where pat_id = " & rEi(i).Pat_id, DBCn, adOpenDynamic, adLockOptimistic
'   If Not rAn.BOF Then
'   If IsNull(rAn!Diabetestyp) Or rAn!Diabetestyp = "" Then
'   If rAna(0).Diabetestyp = "" Then
     If InStr(dt, "estat") > 0 Then fDTyp = "g"
     If InStr(dt, "ekund") > 0 Or InStr(dt, "pankreopri") > 0 Then fDTyp = "s"
     If InStr(dt, "path") > 0 Then fDTyp = "p"
     If InStr("12", Left(dt, 1)) > 0 And Left(dt, 1) <> "" Then fDTyp = dt
     If InStr(dt, "?") > 0 Then fDTyp = fDTyp + "?"
'     erg = Left(erg, rAn.Fields("Diabetestyp").DefinedSize)
'     Call DBCn.Execute("update anamnesebogen set diabetestyp = '" & erg & "' where pat_id = " & rEi(i).Pat_id, rAf)
'     rAna(0).Diabetestyp = erg
'     If rAf <> 1 Then Stop
'   End If
'   End If
  End If
 Next i
End Function ' fDtyp
Function fDMSchL&()
 Dim i&
 fDMSchL = 0
 For i = 1 To UBound(rEi)
  If rEi(i).Inhalt Like "*. sitzung*" Then
   fDMSchL = fDMSchL + 1
  End If
 Next i
End Function ' fDMSchl
Function fFanfFuell()
'FAnfFuell
 Dim lpid&, aktpid&, bm$, lAusgst As Date, i&, j&
 lAusgst = 0
 For i = 1 To UBound(rFa)
  rFa(i).Fanf = 0
  If rFa(i).SchGr = "90" Or rFa(i).SchGr = "41" Then
   rFa(i).Fanf = rFa(i).BhFB
  ElseIf rFa(i).ausgst <> 0 Then
   rFa(i).Fanf = rFa(i).ausgst
  End If
  If rFa(i).Fanf = 0 Then
   If rFa(i).lVorl <> 0 And DateValue(rFa(i).lVorl) <= lAusgst Then
    rFa(i).Fanf = rFa(i).lVorl
   End If
  End If
  If rFa(i).Fanf = 0 Then
   Dim Eintr1 As Date
   Eintr1 = 0
' sql = "select * from Eintraege where pat_id = " + CStr(rFa(i).Pat_ID) + " and zeitpunkt >= datevalue('" + Format(rFa(i).BhFB, "DD.MM.YY") + "') "
   For j = 1 To UBound(rEi)
    If rEi(j).Zeitpunkt > rFa(i).BhFB And (rFa(i).BhFE1 = 0 Or rEi(j).Zeitpunkt <= rFa(i).BhFE1) Then
     If rEi(j).Zeitpunkt < Eintr1 Or Eintr1 = 0 Then
      Eintr1 = rEi(j).Zeitpunkt
     End If
    End If
   Next j
   If Eintr1 <> 0 Then
    If Eintr1 > rFa(i).lVorl And rFa(i).lVorl <> 0 Then
     rFa(i).Fanf = rFa(i).lVorl
    Else
     rFa(i).Fanf = Eintr1
    End If
   Else
    If rFa(i).lVorl <> 0 Then
     rFa(i).Fanf = rFa(i).lVorl
    ElseIf rFa(i).Pat_ID = 2 And rFa(i).Nachname Like "Muster*" Then
     rFa(i).Fanf = #7/1/2004#
    Else
     Print #322, rFa(i).Pat_ID & ": Hatnäckiger Fall bei der Fallanfangsfestlegung"
     rFa(i).Fanf = rFa(i).BhFB
    End If
   End If
  End If
  lAusgst = rFa(i).ausgst
 Next i
End Function ' fFAnfFuell
Function alleSpeichern(frm As Lese)
' rAna!Vorgestellt = MYDAT(Vorgestellt)
 Dim VorgstNeu$, Cpt$, i&, j&
 Dim DTyp$, DMSchL&
 On Error GoTo fehler
 frm.SBez.BackColor = &HFF&
 DoEvents
 DTyp = fDTyp()
 DMSchL = fDMSchL()
 Call fFanfFuell
 
 Call DBCn.BeginTrans
 
' Medklass, redesigned, erster Teil
 If obMySQL And obForeign Then Call DBCn.Execute("set FOREIGN_KEY_CHECKS = 0")
 Call DBCn.Execute("update medarten set pat_id = 0 where pat_id = " & rNa(0).Pat_ID, raf)
 Dim Medi As MediCl
 For i = 1 To UBound(rMe)
  Set Medi = New MediCl
  Medi.Med = Trim(UCase(Left(rMe(i).Medikament, IIf(InStr(rMe(i).Medikament, " ") > 0, InStr(rMe(i).Medikament, " ") - 1, Len(rMe(i).Medikament)))))
  Medi.LT = rMe(i).Medikament
  Medi.Pat_ID = rNa(0).Pat_ID
  Call MedSL.sCAdd(Medi, True)
 Next i
 For i = 1 To UBound(rRe)
  Set Medi = New MediCl
  Medi.Med = Trim(UCase(Left(rRe(i).Medikament, IIf(InStr(Mid(rRe(i).Medikament, 4), " ") > 0, InStr(Mid(rRe(i).Medikament, 4), " ") - 1 + 4, Len(rRe(i).Medikament)))))
  Medi.LT = rRe(i).Medikament
  Medi.Pat_ID = rNa(0).Pat_ID
  Call MedSL.sCAdd(Medi, True)
 Next i

' Call dbcn.Execute("set autocommit = 0", , adAsyncExecute)
' Call dbcn.Execute("start transaction") '("begin")
#If False Then
 If ProgrammLauf(-1, Cpt, -1) Then GoTo Ende
#End If

 If Not rAnaAdo Is Nothing Then If rAnaAdo.State = 1 Then rAnaAdo.Close
 rAnaAdo.CursorLocation = adUseClient
 For i = 1 To 2
  If Not rAnaAdo Is Nothing Then If rAnaAdo.State = 1 Then rAnaAdo.Close
  Call rAnaAdo.Open("select * from anamnesebogen where pat_id = " & rNa(0).Pat_ID, DBCn, adOpenDynamic, adLockOptimistic)
  If rAnaAdo.BOF Then
   Dim primnr&
   If Not rsAdo Is Nothing Then If rsAdo.State = 1 Then rsAdo.Close
   rsAdo.Open "select max(prim) as mprim from anamnesebogen", DBCn, adOpenStatic, adLockReadOnly
   Set rsAdo = DBCn.Execute("select max(prim) as mprim from anamnesebogen")
   If IsNull(rsAdo!mprim) Then primnr = 1 Else primnr = rsAdo!mprim + 1
   Call DBCn.Execute("insert into anamnesebogen (pat_id,prim) values (" & rNa(0).Pat_ID & "," & primnr & ")")
  Else
   Exit For
  End If
 Next i
 If Tkz And Not rAnaAdo!Tkz Then rAnaAdo!Tkz = True
 If rAnaAdo!Nachname <> rNa(0).Nachname Then rAnaAdo!Nachname = rNa(0).Nachname
 If rAnaAdo!Vorname <> rNa(0).Vorname Then rAnaAdo!Vorname = rNa(0).Vorname
 If rAnaAdo!GebDat <> rNa(0).GebDat Then rAnaAdo!GebDat = rNa(0).GebDat
 If rAnaAdo!NVorsatz <> rNa(0).NVorsatz Then rAnaAdo!NVorsatz = rNa(0).NVorsatz
 If rAnaAdo!Titel <> rNa(0).Titel Then rAnaAdo!Titel = rNa(0).Titel
 If rAnaAdo!Vorgestellt <> VorStDat And VorStDat <> 0 Then rAnaAdo!Vorgestellt = VorStDat
 If rAnaAdo!Diabetestyp = "" Then rAnaAdo!Diabetestyp = fDTyp
 If rAnaAdo!DMSchulz <> DMSchulz Then rAnaAdo!DMSchulz = DMSchulz
 If rAnaAdo!RRSchulz <> RRSchulz Then rAnaAdo!RRSchulz = RRSchulz
 If rAnaAdo!DMSchL <> DMSchL Then rAnaAdo!DMSchL = DMSchL
 If rAnaAdo!DMPhier <> DMPDat Then rAnaAdo!DMPhier = DMPDat
 If Not rAnaAdo!obMedNetz And obMedNetz Then rAnaAdo!obMedNetz = obMedNetz
 If Not rAnaAdo!Tkz And Tkz Then rAnaAdo!Tkz = Tkz
' Versicherungsart / Scheingruppe
 If rAnaAdo!Versicherungsart <> rFa(UBound(rFa)).SchGr Then rAnaAdo!Versicherungsart = rFa(UBound(rFa)).SchGr
 Call AnAlle
 Select Case rAnaAdo.EditMode
  Case adEditAdd, adEditInProgress
   On Error Resume Next
   rAnaAdo.Update
   If Err.Number <> 0 Then
    rAnaAdo.CancelUpdate
    Call anaUpd(rAnaAdo)
   End If
   On Error GoTo fehler
 End Select
 
' eintragen
 frm.SBez = "."
 DoEvents
 
 For i = 1 To UBound(rLa)
  If rsAdo Is Nothing Then Else If rsAdo.State = 1 Then rsAdo.Close
  Dim sL1 As sLpar, sL2 As sLpar
  Set sL1 = New sLpar
  sL1.Abkü = rLa(i).Abkü
  sL1.Einheit = rLa(i).Einheit
  sL1.Langtext = rLa(i).Langtext
  Set sL2 = sListLpar.GetItem(sL1)
  If sL2 Is Nothing Then
   Set rsAdo = DBCn.Execute("insert into laborparameter (abkü,einheit,langtext) values('" & rLa(i).Abkü & "','" & rLa(i).Einheit & "','" & rLa(i).Langtext & "')")
'   Set rsAdo = dbcn.Execute("select abkü,eiheit,langtext from laborparameter where abkü = '" & rLa(i).Abkü & "' and einheit = '" & rLa(i).Einheit & "')")
   Call sListLpar.sCAdd(sL1)
  Else
  End If
 Next i
 
 For i = 1 To UBound(rFm)
  If rsAdo Is Nothing Then Else If rsAdo.State = 1 Then rsAdo.Close
  If obVorb Then
   Dim fi1 As sFeldInh, fi2 As sFeldInh
   Set fi1 = New sFeldInh
   fi1.FeldInh = rFm(i).FeldInh
   Set fi2 = sListFldInh.GetItem(fi1)
   If fi2 Is Nothing Then
    Set rsAdo = DBCn.Execute("insert into forminhaltfeldinh (feldinh,stbyte) values('" & rFm(i).FeldInh & "'," & AktByte & ")")
    Set rsAdo = DBCn.Execute("select feldinhvw,feldinh from forminhaltfeldinh where feldinh = '" & rFm(i).FeldInh & "'")
    fi1.FeldInhVW = rsAdo!FeldInhVW
    Call sListFldInh.sCAdd(fi1)
    rFm(i).FeldInhVW = fi1.FeldInhVW
   Else
    rFm(i).FeldInhVW = fi2.FeldInhVW
   End If
  Else
   If obMySQL Then
    For j = 1 To 2
     Set rsAdo = DBCn.Execute("select feldinhvw,feldinh from forminhaltfeldinh where feldinh = '" & rFm(i).FeldInh & "'")
     If rsAdo.BOF Then
      Set rsAdo = DBCn.Execute("insert into forminhaltfeldinh (feldinh,stbyte) values('" & rFm(i).FeldInh & "'," & AktByte & ")", raf)
     Else
      Exit For
     End If
    Next j
   Else
    Dim sqlakt
    sqlakt = "select feldinhvw,feldinh,stbyte from forminhaltfeldinh where feldinh = '" & rFm(i).FeldInh & "'"
    If rsAdo.State = 1 Then Set rsAdo = New ADODB.Recordset
    rsAdo.Open sqlakt, DBCn, adOpenDynamic, adLockOptimistic
    If rsAdo.BOF Then
     rsAdo.AddNew
     rsAdo!FeldInh = rFm(i).FeldInh
     rsAdo!StByte = AktByte
     rsAdo.Update
    End If
   End If
   rFm(i).FeldInhVW = rsAdo!FeldInhVW
  End If
    
  If rsAdo Is Nothing Then Else If rsAdo.State = 1 Then rsAdo.Close
  If obVorb Or 1 = 1 Then
   Dim f1 As sFeld, f2 As sFeld
   Set f1 = New sFeld
   f1.Feld = rFm(i).Feld
   Set f2 = sListFeld.GetItem(f1)
   If f2 Is Nothing Then
    Set rsAdo = DBCn.Execute("insert into forminhaltfeld (Feld,stbyte) values('" & rFm(i).Feld & "'," & AktByte & ")")
    Set rsAdo = DBCn.Execute("select Feldvw,Feld from forminhaltfeld where Feld = '" & rFm(i).Feld & "'")
    f1.FeldVW = rsAdo!FeldVW
    Call sListFeld.sCAdd(f1)
    rFm(i).FeldVW = f1.FeldVW
   Else
    rFm(i).FeldVW = f2.FeldVW
   End If
  Else
   For j = 1 To 2
    Set rsAdo = DBCn.Execute("select feldvw,feld from forminhaltfeld where feld = '" & rFm(i).Feld & "'")
    If rsAdo.BOF Then
     Set rsAdo = DBCn.Execute("insert into forminhaltfeld (feld,stbyte) values('" & rFm(i).Feld & "'," & AktByte & ")")
    Else
     Exit For
    End If
   Next j
   rFm(i).FeldVW = rsAdo!FeldVW
  End If
 Next i
 
 frm.SBez = "S"
 DoEvents
 'rFm.FormID muss nach dem Speichern noch auf rFr.Form_id übertragen werden
 'rFo.FoID muss noch auf rFm.FoID übertragen werden
 ' rFm formulare -> rfmo
 ' rfo forminhkopf -> rfoo
 ' rfr forminhfeld -> rFm
'Public rFo() As formulare
'Public rFr() As forminhkopf
'Public rFm() As forminhfeld

' Anamnesebogenfelder einstellen (-> in Access: Function FehlendeAnamneseBögen())
 For i = 1 To UBound(rDo)
  If InStr(rDo(i).DokName, "namnese") > 0 Then
   If InStr(rDo(i).DokName, "allg") > 0 Then
    rNa(0).AnAllgda = True
   ElseIf InStr(rDo(i).DokName, "2") > 0 Then
    rNa(0).An2da = True
   ElseIf InStr(rDo(i).DokName, "1") > 0 Then
    rNa(0).An1da = True
   End If
  ElseIf InStr(rDo(i).DokName, "heckliste") > 0 Then
   rNa(0).Checkda = True
  End If
 Next i
 
 Call doSpeichern
 If obMySQL And obForeign Then Call DBCn.Execute("set FOREIGN_KEY_CHECKS = 1")
Ende:
' Call dbcn.Execute("commit")
 Call DBCn.CommitTrans
' Call dbcn.Execute("set autocommit = 1")
 frm.SBez.BackColor = &HE0E0E0
 DoEvents
 Exit Function
fehler:
  Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in alleSpeichern/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' alleSpeichern
Function testvor()
 Dim te As Date
 te = BDTtoDate("00000000")
 Call VorstellSetz(te)
End Function ' testvor
Function VorstellSetz(dat) ' Datum setzen, an dem sich der Patient zum ersten Mal vorgestellt hat, 4102, 4109, 3610
 On Error GoTo fehler
 If Not obvorgestellt Then
  VorStDat = dat
  If VorStDat <> 0 Then obvorgestellt = True
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in VorstellSetz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' VorstellSetz(Dat$)
Function aufSplit%(ByVal Q$)
Dim Ind%
On Error GoTo fehler
Ind = 0
ReDim Arra(Ind)
If Not IsNull(Q) Then
 While InStr(Q, "#") > 0
  Arra(Ind) = Left(Q, InStr(Q, "#") - 1)
  Ind = Ind + 1
  ReDim Preserve Arra(Ind)
  Q = Right(Q, Len(Q) - InStr(Q, "#"))
 Wend
 Arra(Ind) = Q
 ArraInd = Ind
End If
aufSplit = Ind
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in aufSplit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' aufSplit

Function ZQuartSort$(Datum As Date) ' Für Abfragen mit Fallzuordnung
Dim j As String * 4, Q As String * 1
On Error GoTo fehler
j = Year(Datum)
Select Case Datum
 Case Is < CDate("1.4." + j): Q = "1"
 Case Is < CDate("1.7." + j): Q = "2"
 Case Is < CDate("1.10." + j): Q = "3"
 Case Else: Q = "4"
End Select
ZQuartSort = j + Q
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZQuartSort/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select ' ZQuartSort

End Function ' ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung

Function ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung
Dim j As String * 4, Q As String * 1
On Error GoTo fehler
j = Year(Datum)
Select Case Datum
 Case Is < CDate("1.4." + j): Q = "1"
 Case Is < CDate("1.7." + j): Q = "2"
 Case Is < CDate("1.10." + j): Q = "3"
 Case Else: Q = "4"
End Select
ZQuart = Q + j
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in ZQuart/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' ZQuart$(Datum As Date) ' Für Abfragen mit Fallzuordnung
Function BDTtoDate(dat) As Date ' für BDT-Format
 On Error GoTo fehler
 If dat = "00000000" Then
 Else
  BDTtoDate = CDate(Format(Left(dat, 8), "##\.##\.####"))
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in BDTtoDate/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' bdttodate
Function BDTtoTime(dat) As Date ' für BDT-Format
 On Error GoTo fehler
 If dat Like "????" Then
  BDTtoTime = CDate(Format(dat, "00:00"))
 ElseIf dat Like "??????" Then
  BDTtoTime = CDate(Format(dat, "00:00:00"))
 ElseIf dat = "*" Then
  BDTtoTime = "00:00"
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in BDTtoTime/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' datForm
Function datForm(dat) ' for vb-Datumsformat oder vb-double (#)
 On Error GoTo fehler
 If IsNull(dat) Then
  datForm = "null"
 ElseIf obMySQL Then
  datForm = "'" + Format(dat, "yyyy-mm-dd hh:mm:ss") + "'"
 Else
  datForm = "#" + Format(dat, "mm\/dd\/yy hh:mm:ss") + "#"
 End If
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in datForm/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' datForm
Function RREintr(RR)
 On Error GoTo fehler
 ReDim Preserve rRr(UBound(rRr) + 1)
 rRr(UBound(rRr)).Pat_ID = rNa(0).Pat_ID
 rRr(UBound(rRr)).Zeitpunkt = messDatum
 rRr(UBound(rRr)).RR = RR
 rRr(UBound(rRr)).absPos = absPos
 rRr(UBound(rRr)).AktZeit = AktZeit
 rRr(UBound(rRr)).FID = rFa(UBound(rFa)).FID
 Exit Function
fehler:
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in datForm/" + App.Path)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function
Function RezEintr(Rez$, obSonder%)
 On Error GoTo fehler
 Dim rRz As New ADODB.Recordset
 If Not IsNull(Rez) Then
  Call aufSplit(Rez)
  ReDim Preserve rRe(UBound(rRe) + 1)
  rRe(UBound(rRe)).Pat_ID = rNa(0).Pat_ID
  rRe(UBound(rRe)).Zeitpunkt = messDatum
  rRe(UBound(rRe)).QS = ZQuartSort(messDatum)
  rRe(UBound(rRe)).QT = ZQuart(messDatum)
  If obSonder Then
   rRe(UBound(rRe)).Rezeptklasse = Arra(2)
   rRe(UBound(rRe)).PZN = Arra(3)
   If ArraInd > 3 Then rRe(UBound(rRe)).Medikament = Arra(4)
  Else
   rRe(UBound(rRe)).Medikament = IIf(ArraInd > 0, Arra(1), "")
   rRe(UBound(rRe)).PZN = IIf(ArraInd > 1, Arra(2), "")
  End If
  rRe(UBound(rRe)).FID = rFa(UBound(rFa)).FID
  rRe(UBound(rRe)).AktZeit = AktZeit
  rRe(UBound(rRe)).Rezept = Arra(0)
  rRe(UBound(rRe)).absPos = absPos
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in RezEintr/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'RezEintr
Function DiagString$(Pat_ID&) ' für dynDiag und tu_brief
 Dim Runde%, rdDi As New ADODB.Recordset, sql$
 On Error GoTo fehler
 sql = "SELECT DiagSicherheit, DiagText, DiagSeite, DiagAttr, ICD, obDauer FROM diagnosen where pat_id = " & CStr(Pat_ID)
' rdDi.Open sql, dbcn, adOpenDynamic, adLockReadOnly
 Set rdDi = DBCn.Execute(sql)
 If rdDi.BOF Then Exit Function
 Runde = 0
 ReDim ICD(0)
 ReDim Diag(0)
 ReDim DiagAttr(0)
 ReDim DiagSe(0)
 ReDim DiagSi(0)
 ReDim gk(0)
 ReDim obDauer(0)
 Do While Not rdDi.EOF
  If Not rdDi!ICD Like "Z25*" Then
   ICD(Runde) = nz(rdDi!ICD,"")
   Diag(Runde) = nz(rdDi!DiagText,"")
   DiagAttr(Runde) = nz(rdDi!DiagAttr,"")
   DiagSe(Runde) = nz(rdDi!DiagSeite,"")
   DiagSi(Runde) = nz(rdDi!DiagSicherheit,"")
   gk(Runde) = False
   obDauer(Runde) = nz(rdDi!obDauer,"")
   Runde = Runde + 1
   ReDim Preserve ICD(Runde)
   ReDim Preserve Diag(Runde)
   ReDim Preserve DiagAttr(Runde)
   ReDim Preserve DiagSe(Runde)
   ReDim Preserve DiagSi(Runde)
   ReDim Preserve gk(Runde)
   ReDim Preserve obDauer(Runde)
  End If
  rdDi.MoveNext
 Loop
 DiagNr = Runde
 DiagString = MachDiagnosen(Pat_ID)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in DiagString/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' DiagString$(Pat_id&)

Function MachDiagnosen$(Pat_ID&) ' für AnaEintragen, MachSammelTab und DiagString
   Dim j&, K&, Runde%, obkNeph%
   On Error GoTo fehler
   obkNeph = obKeineNephropathie(Pat_ID)
    For j = 0 To DiagNr - 1
     If Not IsNull(ICD(j)) And ICD(j) <> "" Then
     If Left(ICD(j), 2) = "E1" Then
      Select Case Mid(ICD(j), 3, 1)
       Case "0"
        Diag(j) = "Diabetes mellitus Typ 1"
       Case "1"
        Diag(j) = "Diabetes mellitus Typ 2"
       Case "2"
        Diag(j) = "Diabetes mellitus in Verbindung mit Fehl- oder Mangelernährung"
       Case "3"
        Diag(j) = "Diabetes mellitus (sekundär)"
       Case "4"
        Diag(j) = "Diabetes mellitus"
      End Select
     End If
     If Left(ICD(j), 5) = "N08.3" Then
      If InStr(Diag(j), "Glomeruläre Krankheiten") > 0 Then
       Diag(j) = "Diabetische Nephropathie"
      End If
     End If
     If InStr(Diag(j), "Niereninsuff.") > 0 Then Diag(j) = Replace(Diag(j), "Niereninsuff.", "Niereninsuffizienz")
     If Left(ICD(j), 5) = "G99.0" Then
      If InStr(Diag(j), "Autonome Neuropathie bei endokrinen") > 0 Then
       Diag(j) = "Diabetische autonome Neuropathie"
      End If
     End If
     If Left(ICD(j), 5) = "I79.9" Or Left(ICD(j), 5) = "I79.2" Then
      If InStr(Diag(j), "Periphere Angiopathie bei anderenorts") > 0 Then
       Diag(j) = "Periphere Angiopathie"
      End If
     End If
' M14.2 = Diabetische Arthropathie, schon richtig
     If Left(ICD(j), 5) = "M14.6" Then
      If InStr(Diag(j), "Neuropathische Arthropathie ") > 0 Then
       Diag(j) = "Diabetische Osteoarthropathie" ' "Charcot-Fuß"
      End If
     End If
     If Left(ICD(j), 5) = "M36.8" Then
      If InStr(Diag(j), "anderenorts") > 0 Then
       Diag(j) = "Diabetische Bindegewebserkrankung"
      End If
     End If
     If Left(ICD(j), 3) = "E66" Then
      Diag(j) = "Übergewicht"
     End If
     If InStr("K76.0 K76.9 K71.6 K71.7 K77.8", Left(ICD(j), 5)) > 0 _
      And InStr(Diag(j), "nbek") = 0 And InStr(Diag(j), "nklar") = 0 And InStr(DiagAttr(j), "nklar") = 0 Then
'     If Left(ICD(j), 5) = "K77.8" And _
       InStr(Diag(j), "Leberkrankheiten bei sonstigen anderenorts klassifizierten Krankheiten") > 0 Then
       Diag(j) = "Nicht-alkoholische Steatohepatose ('NASH-Syndrom')"
     End If
     If Left(ICD(j), 3) = "L89" Then
      Diag(j) = "Diabetisches Fußsyndrom "
      If Mid(ICD(j), 5, 1) < "9" Then _
       Diag(j) = Diag(j) + " im Stadium Wagner "
      Select Case Mid(ICD(j), 5, 1)
       Case "1": Diag(j) = Diag(j) + "0"
       Case "2": Diag(j) = Diag(j) + "1"
       Case "3": Diag(j) = Diag(j) + "2"
       Case "4": Diag(j) = Diag(j) + "3"
      End Select
     End If
'
'     If Left(ICD(j), 5) = "H28.0" Then Diag(j) = " Diabetische Katarakt"
'     If Left(ICD(j), 5) = "H36.0" Then Diag(j) = " Diabetische Retinopathie"
'     If Left(ICD(j), 5) = "G63.2" Then Diag(j) = " Diabetische Polyneuropathie"
     Select Case DiagSi(j)
      Case "V"
       If Left(Diag(j), 4) <> "V.a." Then Diag(j) = "V.a. " + Replace(Replace(Replace(Replace(Diag(j), "Diab.", "diab."), "Diabeti", "diabeti"), "Diabet.", "diabet."), "Auton", "auton")
      Case "Z"
       If Left(Diag(j), 4) <> "Z.n." Then Diag(j) = "Z.n. " + Replace(Replace(Replace(Replace(Diag(j), "Diab.", "diab."), "Diabeti", "diabeti"), "Diabet.", "diabet."), "Auton", "auton")
      Case "A"
       If Left(Diag(j), 9) <> "Ausschluß" Then Diag(j) = "Ausschluß " + Replace(Replace(Replace(Replace(Diag(j), "Diab.", "diab."), "Diabeti", "diabeti"), "Diabet.", "diabet."), "Auton", "auton")
     End Select
     Select Case DiagSe(j)
      Case "R"
       If Right(Diag(j), 3) <> " re" Then Diag(j) = Diag(j) + " re"
      Case "L"
       If Right(Diag(j), 3) <> " li" Then Diag(j) = Diag(j) + " li"
      Case "B"
       If InStr(Diag(j), "bds.") = 0 Then Diag(j) = Diag(j) + " bds."
     End Select
     If DiagAttr(j) <> "" Then Diag(j) = Diag(j) + " (" + DiagAttr(j) + ")"
     Diag(j) = Replace(Diag(j), ", nicht näher bezeichnet", "")
     End If
    Next j
    For j = 0 To DiagNr - 1
     For K = j + 1 To DiagNr - 1
      If Diag(j) = Diag(K) And ICD(j) = ICD(K) Then
       Diag(K) = ""
       ICD(K) = ""
      End If
     Next K
    Next j
   MachDiagnosen = ""
   Dim ohneNotwend%
   ohneNotwend = True
   For Runde = 0 To 17
    For j = 0 To DiagNr - 1
     If Not IsNull(ICD(j)) And ICD(j) <> "" Then
     If (Diag(j) <> "" Or ICD(j) <> "") And _
        (Not ohneNotwend Or InStr(Diag(j), "otwend") = 0) And _
        (Not obkNeph Or InStr(ICD(j), "N08.3") = 0) _
        Then
      If (Runde = 0 And (Left(ICD(j), 2) = "E1" Or InStr("O24.4 R73.0", Left(ICD(j), 5)) > 0)) Or _
        (Not gk(j) And Runde = 1 And (InStr("N08.3", Left(ICD(j), 5)) > 0)) Or _
        (Not gk(j) And Runde = 2 And (InStr("N17 N18 N19", Left(ICD(j), 3)) > 0)) Or _
        (Not gk(j) And Runde = 3 And (InStr("G59 G62", Left(ICD(j), 3)) <> 0 Or _
             InStr("H28.0 H36.0 G63.2 G99.0 I79.2 K77.8 K76.0 K76.9 K71.7 K71.6", Left(ICD(j), 5)) <> 0 Or _
             InStr("L89. I70.", Left(ICD(j), 4)) > 0)) Or _
        (Not gk(j) And Runde = 4 And (Left(ICD(j), 3) = "I10" Or Left(ICD(j), 3) = "I15")) Or _
        (Not gk(j) And Runde = 5 And (Left(ICD(j), 2) = "I1" Or ICD(j) = "H35.0")) Or _
        (Not gk(j) And Runde = 6 And _
             (InStr("E65 E66 E67 E68", Left(ICD(j), 3)) > 0 Or InStr("E78. F17.", Left(ICD(j), 4)) > 0)) Or _
        (Not gk(j) And Runde = 7 And _
             (InStr("I25. I20. I21. I50.", Left(ICD(j), 4)) > 0 Or InStr("Z95.1 Z92.1", Left(ICD(j), 5)) > 0)) Or _
        (Not gk(j) And Runde = 8 And (InStr("R60 R01", Left(ICD(j), 3)) > 0) Or InStr("I83.1", Left(ICD(j), 4) > 0)) Or _
        (Not gk(j) And Runde = 9 And (InStr("I63 I64", Left(ICD(j), 3)) > 0)) Or _
        (Not gk(j) And Runde = 10 And (InStr("I7", Left(ICD(j), 2)) > 0 Or InStr("Z95.88", Left(ICD(j), 6)) > 0)) Or _
        (Not gk(j) And Runde = 11 And (InStr("M36.8 T87.4", Left(ICD(j), 5)) <> 0 Or _
             InStr("M14. T89. T79.", Left(ICD(j), 4)) > 0)) Or _
        (Not gk(j) And Runde = 12 And obDauer(j) And DiagSi(j) = "G") Or _
        (Not gk(j) And Runde = 13 And obDauer(j) And DiagSi(j) = "V") Or _
        (Not gk(j) And Runde = 14 And Not obDauer(j) And DiagSi(j) = "G") Or _
        (Not gk(j) And Runde = 15 And Not obDauer(j) And DiagSi(j) = "V") Or _
        (Not gk(j) And Runde = 16 And DiagSi(j) = "Z") Or _
        (Not gk(j) And Runde = 17 And DiagSi(j) = "A") Then
       MachDiagnosen = MachDiagnosen + IIf(Runde > 0 And Runde < 4, " ", "") + Diag(j) + Chr(9) + "[" + ICD(j) + "]" + Chr(11)
       gk(j) = True ' abgehakt
      End If ' (Runde = 0 And (Left(icd(j), 2) = "E1") Or InStr("O24.4 R73.0", Left(icd(j), 5) > 0)) Or
     End If ' Diag(j) <> "" Or icd(j) <> "" Then
     End If
    Next ' j = 0 To DiagNr - 1
   Next ' Runde = 0 To 7
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in MachDiagnosen/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' machDiagnosen$()


Function obKeineNephropathie%(Pat_ID&, Optional obMakroAlb%)
 Dim lAlbS$, lKreS$, rsAdo As New ADODB.Recordset, sql$
 On Error GoTo fehler
 obMakroAlb = 0
 On Error Resume Next
 sql = "select Pat_ID, ZeitPunkt, FertigStGrad, AbKü, LangText,Wert, Einheit, Kommentar,"""" as NB from labor where pat_id = " + CStr(Pat_ID) + " UNION SELECT Pat_ID, Eingang AS zeitpunkt, BefArt AS FertigStGrad, Abkü, langname AS Langtext, Wert, Einheit, Kommentar,Normbereich as NB " + _
  "FROM LaborXUS LEFT JOIN LaborXWert ON [LaborXUS].[RefNr]=[LaborXWert].[RefNr] " + _
  "WHERE pat_id = " + CStr(Pat_ID) + " and not exists (select * from laborNeu where pat_id = " + CStr(Pat_ID) + " and abkü = LaborXWert!Abkü and wert = LaborXWert.wert and zeitpunkt > LaborXUS.Eingang -3 and zeitpunkt < LaborXUS.Eingang+6)"
 'lAlbS = Dtb.OpenRecordset("select nz(wert,iif(isnull(kommentar),"""",kommentar)) as erg from (" + sql + ") where abkü = ""ALBCRE"" order by zeitpunkt desc")!erg
' rsAdo.Open "select nz(wert,iif(isnull(kommentar),"""",kommentar)) as erg from (" + sql + ") where abkü = ""ALBCRE"" order by zeitpunkt desc", dbcn, adOpenStatic, adLockReadOnly
 Set rsAdo = DBCn.Execute("select nz(wert,iif(isnull(kommentar),"""",kommentar)) as erg from (" + sql + ") where abkü = ""ALBCRE"" order by zeitpunkt desc")
 If Not rsAdo.EOF Then lAlbS = rsAdo!erg
 On Error GoTo fehler
 If lAlbS = "" Then Exit Function
 If InStr(lAlbS, "<") = 0 Then
  lAlbS = Replace(lAlbS, ".", ",")
  If Not IsNumeric(lAlbS) Then Exit Function
  If CDbl(lAlbS) >= 200 Then obMakroAlb = True
  If CDbl(lAlbS) >= 10 Then Exit Function ' setzen wir mal 10 als Grenze an
 End If
 On Error Resume Next
 rsAdo.Close
' rsAdo.Open "select nz(wert,iif(isnull(kommentar),"""",kommentar)) as erg from (" + sql + ") where abkü = ""KREA"" order by zeitpunkt desc", dbcn, adOpenStatic, adLockReadOnly
 Set rsAdo = DBCn.Execute("select nz(wert,iif(isnull(kommentar),"""",kommentar)) as erg from (" + sql + ") where abkü = ""KREA"" order by zeitpunkt desc")
 If Not rsAdo.EOF Then lKreS = rsAdo!erg
' lKreS = Dtb.OpenRecordset("select nz(wert,iif(isnull(kommentar),"""",kommentar)) as erg from (" + sql + ") where abkü = ""KREA"" order by zeitpunkt desc")!erg
 On Error GoTo fehler
 If IsNumeric(lKreS) And lKreS <> 0 Then
  If CDbl(Replace(lKreS, ".", ",")) >= 1.3 Then
   Exit Function
  End If
  obKeineNephropathie = -1
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in obKeineNephropathie/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' obKeineNephropathie%(Pat_id&, Optional obMakroAlb%)

Public Function GesName$(NVorsatz$, Nachname$, Titel, Vorname$, Optional GebDat As Date)  ' s.a. Gesnam(
 On Error GoTo fehler
 GesName = NVorsatz + IIf(NVorsatz = "", "", " ") + Nachname + ", " + IIf(IsNull(Titel) Or Titel = "", "", Titel + ", ") + Vorname
 If GebDat > 0 Then GesName = GesName + ", *" + Format(GebDat, "dd/mm/yy")
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in GesName/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' GesName

Function PatInit()
' Pat_id = 0
 On Error GoTo fehler
 NEin0 = 0
' lfdPatNr = 0
   lRR = ""
   gesRR = "" ' RR aus Turbomed
'   Versicherung = ""
   AlbErhöht = ""
   AlbDat = 0
   lHbA1c = ""
   lHbA1cDat = 0
   Arr = "" ' Aktueller Blutdruck
   obMedPlan = False
   obMedPlanGelesen = False
   medplan = ""
   obvorgestellt = False
' Diagnosedaten
 DiagNr = 0
' Anamnesebogen
 DMSchulz = 0
 RRSchulz = 0
 DMSchL = 0
 Tkz = 0
 VorStDat = 0
 obMedNetz = 0
 
 RRSchulz = 0
 DMPDat = 0 'Null
' FID = 0
 ReDim Diag(DiagMaxZahl)
 ReDim ICD(DiagMaxZahl)
 ReDim DiagSi(DiagMaxZahl)
 ReDim DiagSe(DiagMaxZahl)
 ReDim DiagAttr(DiagMaxZahl)
 ReDim diagDat(DiagMaxZahl)
 ReDim obDauer(DiagMaxZahl)
 ReDim gk(DiagMaxZahl)
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in PatInit/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' PatInit()
Function Stutz$(Q$)
 On Error GoTo fehler
    If Right(Q, 1) = "^" Then
     Stutz = Left(Q, Len(Q) - 1)
    Else
     Stutz = Q
    End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in Stutz/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function 'Stutz

Function PraxisHbA1c(EintrArt$, EintrInh$)
 Dim Ep%, ZwS$, erg$
 If UCase(EintrArt) = "HBA1C" Then
  FStG = "E"
  Abkü = "HBA1C"
  Einheit = "%"
  
  Ep = InStr(EintrInh, ",")
  If Ep > 1 Then
   ZwS = Left(EintrInh, Ep - 1)
   erg = ""
   If IsNumeric(Right(ZwS, 2)) Then
    erg = Right(ZwS, 2)
   ElseIf IsNumeric(Right(ZwS, 1)) Then
    erg = Right(ZwS, 1)
   End If
   If erg <> "" Then
    ZwS = Mid(EintrInh, Ep + 1, 1)
    If IsNumeric(ZwS) Then
     erg = erg + "," + ZwS
    End If
   End If
  Else
   Ep = InStr(EintrInh, "%")
   If Ep > 0 Then
    ZwS = Trim(Left(EintrInh, Ep - 1))
    If IsNumeric(Right(ZwS, 2)) Then
     erg = Right(ZwS, 2)
    ElseIf IsNumeric(Right(ZwS, 1)) Then
     erg = Right(ZwS, 1)
    End If
   End If
  End If
  If IsNumeric(erg) Then
   Wert = erg
   Langtext = "HbA1c"
   Einheit = "%"
   Call LaborEintr0
   Kommentar = "Praxis-HbA1c"
   Call Laboreintr1
  End If
 End If
End Function ' PraxisHbA1c
Function LTEinfüg&(Langtext$)
  Dim i&
  On Error GoTo fehler
  If Langtext = lLang Then
   LTEinfüg = lLangVW
  Else
   For i = 1 To 2
    If Not rsAdo Is Nothing Then If rsAdo.State = 1 Then rsAdo.Close
'   rsAdo.Open "select langtext,langtextvw from laborlangtext where langtext = '" & Langtext & "'", dbcn, adOpenStatic, adLockReadOnly
    Set rsAdo = DBCn.Execute("select langtext,langtextvw from laborlangtext where langtext = '" & Langtext & "'")
    If rsAdo.EOF Then
     Set rsAdo = DBCn.Execute("insert into laborlangtext (langtext) values ('" & Langtext & "')")
    Else
     Exit For
    End If
   Next i
   LTEinfüg = rsAdo!LangtextVW
   lLangVW = LTEinfüg
   lLang = Langtext
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LTEinfüg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LTEinfüg&(Langtext$)
Function KomEinfüg&(Kommentar$)
  Dim i&
  On Error GoTo fehler
  If Kommentar = lKomm Then
   KomEinfüg = lKommVW
  Else
   For i = 1 To 2
    Set rsAdo = Nothing
'   rsAdo.Open "select Kommentar,Kommentarvw from laborkommentar where Kommentar = '" & Kommentar & "'", dbcn, adOpenStatic, adLockReadOnly
    Set rsAdo = DBCn.Execute("select Kommentar,Kommentarvw from laborkommentar where Kommentar = '" & Kommentar & "'")
    If rsAdo.EOF Then
     Set rsAdo = DBCn.Execute("insert into laborkommentar (Kommentar) values ('" & Kommentar & "')")
    Else
     Exit For
    End If
   Next i
   KomEinfüg = rsAdo!KommentarVW
   lKomm = Kommentar
   lKommVW = KomEinfüg
  End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LTEinfüg/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function

Function LaborEintr0()
 On Error GoTo fehler
 Dim obneu%, obüber%, i%
 Dim rsAdo As New ADODB.Recordset
' Abkü = nz(Abkü,"") 'If IsNull(rLab!Abkü) Then rLab!Abkü = ""
' Einheit = ZeichenSatz(nz(Einheit,"")) ' If IsNull(rLab!Einheit) Then rLab!Einheit = ""
 If Einheit = "" Then Einheit = "kA" ' 2.10.: ist so in Turbomed, => auch für die Direktübernahme so hindrehen
#If True Then
 obneu = -1
 For i = 1 To UBound(rLa)
  If rLa(i).Zeitpunkt = messDatum And rLa(i).Abkü = Abkü Then
   If rLa(i).FertigStGrad = "T" And FStG = "E" Then
    obüber = -1: obneu = 0: ls = i
   ElseIf FStG = "T" Then
    obneu = 0
   End If
   Exit For
  End If
 Next i
 If obneu Or obüber Then
  If obneu Then ReDim Preserve rLa(UBound(rLa) + 1): ls = UBound(rLa)
  rLa(ls).Pat_ID = rNa(0).Pat_ID
  rLa(ls).Zeitpunkt = messDatum
  rLa(ls).FertigStGrad = FStG
  rLa(ls).Abkü = Abkü
  rLa(ls).Einheit = Einheit
  rLa(ls).Pat_ID = rNa(0).Pat_ID
  rLa(ls).Wert = Wert
  rLa(ls).AktZeit = AktZeit
  rLa(ls).FID = rFa(UBound(rFa)).FID
  rLa(ls).absPos = absPos
  rLa(ls).Langtext = Langtext
'  If Langtext <> "" And Not IsNull(Langtext) Then
  rLa(ls).LangtextVW = LTEinfüg&(Langtext)
'  End If
'  If rLa(ls).Kommentar <> "" And Not IsNull(rLa(ls).Kommentar) Then
   Kommentar = ""
   rLa(ls).KommentarVW = KomEinfüg&(Kommentar)
'  End If
 End If
 Langtext = ""
' Kommentar = ""
#Else
  Dim nmE%, nmT%, LTVW$
  Dim neuWert$
  rsAdo.Open "select * from laborneu where pat_id = " & rNa(0).Pat_ID & " and zeitpunkt = " & datForm(messDatum) & " and fertigstgrad = 'T' and abkü = '" & Abkü & "'", DBCn, adOpenStatic, adLockReadOnly
  nmT = rsAdo.EOF
  Select Case FStG
   Case "E", ""  ' t suchen, falls gefunden, überschreiben, falls nicht, neu anlegen
     If nmT Then obneu = -1 Else obüber = -1
   Case "T"   ' beides suchen, falls gefunden, weglassen, falls nicht, neu anlegen
     If nmT Then
      rsAdo.Close
      rsAdo.Open "select * from laborneu where pat_id = " & rNa(0).Pat_ID & " and zeitpunkt = " & datForm(messDatum) & " and fertigstgrad = 'E' and abkü = '" & Abkü & "'", DBCn, adOpenStatic, adLockReadOnly
      obneu = rsAdo.EOF
     End If
  End Select
   ' wenn der vermutlich zugehörige Teilbefund nach einem Endbefund eingelesen wird, dann lieber nichts tun
' das im DAO-Text weiter unten stehende kommt hier vorher
  If obneu Or (obüber And Wert <> "") Then
   For i = 1 To 2
    rsAdo.Close
    rsAdo.Open "select langtext,langtextvw from laborlangtext where langtext = '" & Langtext & "'", DBCn, adOpenStatic, adLockReadOnly
    If rsAdo.EOF Then
     Set rsAdo = DBCn.Execute("insert into laborlangtext (langtext) values '" & Langtext & "'")
    Else
     Exit For
    End If
   Next i
   LTVW = rsAdo!LangtextVW
  End If
  If obüber Then
   If Wert <> "" Then
    Set rsAdo = DBCn.Execute("update laborneu set wert = '" & neuWert & "', einheit = '" & Einheit & "', langtextvw = '" & LTVW & "', fid = " & rFa(UBound(rFa)).FID & ", abspos = " & absPos & ", stbyte = " & AktByte & ", aktzeit = " & datForm(AktZeit) & " where pat_id = " & rNa(0).Pat_ID & " and zeitpunkt = " & datForm(messDatum) & " and fertigstgrad = 'E' and abkü = '" & Abkü & "'", numA)
    If numA <> 1 Then Stop
   End If
  ElseIf obneu Then
   If obMySQL And obForeign Then Call DBCn.Execute("set foreign_key_checks=0")
   Set rsAdo = DBCn.Execute("insert into laborneu (pat_id,zeitpunkt,fertigstgrad,abkü,wert,einheit,langtextvw,fid,abspos,aktzeit,stbyte) values (" & rNa(0).Pat_ID & "," & datForm(messDatum) & ",'" & FStG & "','" & Abkü & "','" & Wert & "','" & Einheit & "','" & LTVW & "'," & rFa(UBound(rFa)).FID & "," & absPos & "," & datForm(AktZeit) & "," & AktByte & ")")
  Else
   Exit Function
  End If
'   Call LaborParameter(Abkü, Einheit, LTVW, Langtext)
'  rLaN.Update
#End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborEintr0/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LaborEintr0()
Function Laboreintr1()
 On Error GoTo fehler
#If True Then
' If Langtext <> "" And Not IsNull(Langtext) Then
'  rLa(ls).LangtextVW = LTEinfüg&(Langtext)
' End If
 If Kommentar <> "" And Not IsNull(Kommentar) Then
  rLa(ls).KommentarVW = KomEinfüg&(Kommentar)
  Kommentar = ""
  rLa(ls).Wert = Wert
 End If
' Langtext = ""
#Else
 Dim rsAdo As New ADODB.Recordset, i%
 For i = 1 To 2
  rsAdo.Open "select kommentarvw,kommentar from laborkommentar where kommentar = '" & Kommentar & "'", DBCn, adOpenStatic, adLockReadOnly
  If rsAdo.EOF Then
   Set rsAdo = DBCn.Execute("insert into laborkommentar (kommentar) values ('" & Kommentar & "')")
  Else
   Exit For
  End If
 Next
 Set rsAdo = DBCn.Execute("update laborneu set wert = '" & Wert & "', kommentarvw = '" & rsAdo!KommentarVW & "' where pat_id = " & rNa(0).Pat_ID & " and zeitpunkt = " & datForm(messDatum) & " and abkü = '" & Abkü & "' and (isnull(wert) or wert = '') and (isnull(kommentarvw) or kommentarvw = '')", numA)
#End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborEintr1/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LaborEintr1()
#If False Then
Function AlleLaborParameter()
 Dim rLab As DAO.Recordset
 Dim rLaP As DAO.Recordset, rLN As DAO.Recordset, Fld As DAO.Field
  On Error GoTo fehler
 Call dtbInit
 Set rLab = Dtb.OpenRecordset("LaborUnion")
 #If False Then
  Set rLaP = Dtb.OpenRecordset("LaborParameter")
  Set rLN = Dtb.OpenRecordset("LaborParameterNeu")
  Do While Not rLaP.EOF
   If Not (rLaP!gruppe = 0 And rLaP!reihe = 0 And IsNull(rLaP!unm) And IsNull(rLaP!onm) And IsNull(rLaP!unw) And IsNull(rLaP!onw)) Then
   ' rLaP.Delete
    rLN.AddNew
    For Each Fld In rLaP.Fields
     rLN.Fields(Fld.Name) = Fld.Value
    Next Fld
    rLN.Update
   End If
   rLaP.Move 1
  Loop
 #End If
 rLab.MoveFirst
 Do While Not rLab.EOF
  Call LaborParameter(rLab)
  rLab.MoveNext
 Loop
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in AlleLaborParameter/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' AlleLaborParameter

Function LaborParameter(rLab As DAO.Recordset)
'Dim rLP As DAO.Recordset, rlaL As DAO.Recordset
Dim Abkü$, Einheit$, Langtext$
Dim obltvw%
On Error Resume Next
obltvw = IIf(Err.Number = 0, -1, 0)
On Error GoTo fehler
Abkü = IIf(IsNull(rLab!Abkü), "", Trim(rLab!Abkü)) 'If IsNull(rLab!Abkü) Then rLab!Abkü = ""
Einheit = Zeichensatz(IIf(IsNull(rLab!Einheit), "", Trim(rLab!Einheit))) ' If IsNull(rLab!Einheit) Then rLab!Einheit = ""
If Einheit = "" Then Einheit = "kA" ' 2.10.: ist so in Turbomed, => auch für die Direktübernahme so hindrehen
Langtext = ""
If obltvw Then
 Set rlaL = TabÖff("LaborLangtext", "LangtextVW")
 If Not rlaL Is Nothing And Not IsNull(rLab!LangtextVW) Then
  rlaL.Seek "=", rLab!LangtextVW
  If Not rlaL.NoMatch Then
   Langtext = rlaL!Langtext
  End If
 End If
Else
 Langtext = nz(rLab!Langtext,"")
End If
Set rLP = TabÖff("LaborParameter", "Abkü")
If Abkü <> "" Then
 rLP.Seek "=", Abkü, Einheit
 If rLP.BOF Then
  rLP.AddNew
 ElseIf rLP.NoMatch Then
  rLP.AddNew
 End If
 If rLP.EditMode = 2 Then
  Print #322, "Neuer Laborparameter: " + Abkü + " [" + Einheit + "]"
  rLP!Abkü = Abkü
  rLP!Einheit = Einheit
 ElseIf (rLP!Langtext <> Langtext) Then 'Or (rLP!Kommentar <> rLab!Kommentar And Not IsNull(rLab!Kommentar)) Then
  rLP.Edit
 Else
  rLPbm = rLP.Bookmark
  GoTo fertig:
 End If
 If rLP.EditMode = 1 Or rLP.EditMode = 2 Then
  If Not IsNull(Langtext) And Langtext <> "" Then
   rLP!Langtext = Langtext
  End If
'  rLP!Kommentar = rLab!Kommentar
  rLP!AktZeit = AktZeit
  rLP.Update
  rLPbm = rLP.LastModified
 End If
End If
 GoTo fertig
test:
 MsgBox "Hier Fehler bei Einheit"
 Resume
fertig:
 rLP.Close
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborParameter/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LaborParameter(rLab as dao.recordset)
#End If

Function formInhMach()
 Stop
 Dim Nr&
 Dim altPat_id&, altZeitpunkt#, altform_id&
 Dim myA As New ADODB.Connection, fn As ADODB.Recordset, fk As New ADODB.Recordset, ff As New ADODB.Recordset
 myA.Open CStrAcc + aVerz & "quelle.mdb"
 myA.Execute ("delete from FormInhFeld")
 myA.Execute ("delete from FormInhKopf")
' myA.Execute ("alter table forminhkopf alter column foid counter (1,1)")
 fk.Open "FormInhKopf", myA, adOpenDynamic, adLockOptimistic
 ff.Open "FormInhFeld", myA, adOpenDynamic, adLockOptimistic
 Set fn = myA.Execute("select * from forminhaltn order by pat_id, Form_ID, FeldVW, ZeitPunkt, FeldNr")
 Do While Not fn.EOF
  If fn!Pat_ID <> altPat_id Or fn!Zeitpunkt <> altZeitpunkt Or altform_id <> fn!Form_ID Then
   fk.AddNew
   fk!FID = fn!FID
   fk!Pat_ID = fn!Pat_ID
   fk!Form_ID = fn!Form_ID
   fk!Form_AbkVW = fn!Form_AbkVW
   fk!Zeitpunkt = fn!Zeitpunkt
   fk!absPos = fn!absPos
   fk!AktZeit = fn!AktZeit
   fk!StByte = fn!StByte
   fk.Update
   altPat_id = fn!Pat_ID
   altZeitpunkt = fn!Zeitpunkt
   altform_id = fn!Form_ID
   Nr = 1
  Else
   Nr = Nr + 1
  End If
  ff.AddNew
  ff!FeldNr = fn!FeldNr
  ff!FeldVW = fn!FeldVW
  ff!FeldInhVW = fn!FeldInhVW
  ff!FoID = fk!FoID
  ff!Nr = Nr
  ff.Update
  fn.MoveNext
 Loop
' Call dbcn.Execute("set foreign_key_checks=0;")
' Call dbcn.Execute("insert into forminh select * from forminhv;")
' Call dbcn.Execute("set foreign_key_checks=1;")
End Function ' formInhMach

Function doQuelldatum(DokName$) As Date
 Dim Spli$(), i&, buch$, cont As String * 10, auffuell%, auffStand%
  Spli = Split(Replace(Replace(Replace(DokName, ".jpg", ""), ".tif", ""), ",", " "))
  For i = UBound(Spli) To 0 Step -1
   If IsDate(Spli(i)) Then
'    rdo.Edit
    doQuelldatum = CDate(Spli(i))
'    rdo.Update
    Exit For
   End If
  Next i
  ReDim Spli(0)
End Function ' doQuelldatum
' 4.9.06: Wird in NachArbeit aufgerufen
Public Function LaborDirektImport(frm As Lese, Optional oballe% = 0)
 Dim abPos&, i&, j&, ii&
 Const debugohneSpeichern% = 0
 Const LabTransPfad$ = "\\server\BioWinBACKUP"
' Const neuVerz$ = "P:\Biowin\Backup\neu\"
 Const ZS As Boolean = True
 Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
 Dim repl$
 Dim fi As Files, Fld As Folder, f1 As File, neuPfad$, keimz%
 Dim RefNr&
 Dim obB% ' obBakt: -1 = ja
 Dim altAusgabe$, txt$, SatzlängenArt%, obSonderSatz%
 Dim DebugZähler&
 Dim aktDateiNr&, DateiZahl&, obDateiNeu%
 Dim merkDatID&
 
 t1a = Now
 altAusgabe = frm.Ausgabe
 frm.ZeilenBez = "Bytes:"
 frm.Ausgabe = "Lese Labor-DFÜ-Dateien ein ..." + vbCrLf + altAusgabe
 On Error GoTo fehler
 If oballe Then 'Reihenfolge wichtig:
  Call LabLösch(frm)
  Call LöschRefNr
 End If
 If FSO Is Nothing Then Set FSO = New FileSystemObject
' Call VerzPrüf(neuVerz)
' If 1 = 1 Then
 Dim SL As New SortierListe, SD As SortierDatei
 Set Fld = FSO.GetFolder(LabTransPfad)
 Set fi = Fld.Files
 frm.BytesBez = "Dateien:"
 frm.GesBytes = fi.Count
 altAusgabe = frm.Ausgabe
 frm.Ausgabe = vbCrLf + "Untersuche Dateien in Verzeichnis " & Fld & " ..." & vbCrLf & frm.Ausgabe
 For Each f1 In fi
  If f1.Name Like "1b*.ld*" Or f1.Name Like "Labor*.dat" Then
    aktDateiNr = aktDateiNr + 1
    frm.Bytes = aktDateiNr
'    repl = f1.Name
'    neuPfad = f1.Path
    obDateiNeu = 0
    For i = 1 To 2
     Set rs = DBCn.Execute("select * from laborxeingel where name = '" & f1.Name & "' and pfad = '" & umw(f1.Path) & "'")
     If rs.BOF Then
      obDateiNeu = True
     Else
      If debugohneSpeichern Then merkDatID = rs!DatID
      Exit For
     End If
    Next i
    If debugohneSpeichern Then obDateiNeu = -1
    frm.Sekunden = Format((Now - t1) * 60 * 60 * 24, "###,###,###,###,###,###,##0")
    If obDateiNeu Then
      frm.Ausgabe = "." + frm.Ausgabe
      If aktDateiNr Mod 100 = 0 Then
       frm.Ausgabe = vbCrLf + frm.Ausgabe
      End If
      DoEvents
      Set SD = New SortierDatei
      If merkDatID <> 0 Then
       SD.DatID = merkDatID
       merkDatID = 0
      End If
      Set SD.File = f1
'      SD.DatID = rs!DatID ' für Debugging
      Call SL.sCAdd(SD)
    End If
  End If
 Next f1
 DateiZahl = SL.Count
 aktDateiNr = 0
 frm.BytesBez = "Dateien:"
 frm.GesBytes = DateiZahl
 frm.Ausgabe = altAusgabe
 frm.Bytes = 0
 For ii = 1 To SL.Count
    aktDateiNr = aktDateiNr + 1
    frm.Ausgabe = "Lese LaborDatei " + CStr(aktDateiNr) + "/" + CStr(DateiZahl) + " (" + SL.Item(ii).File.Path + ") ein ..." + vbCrLf + frm.Ausgabe
    If SL.Item(ii).DatID = 0 Then
      Set rs = DBCn.Execute("insert into laborxeingel (name,pfad,zp) values ('" & SL.Item(ii).File.Name & "','" & umw(SL.Item(ii).File.Path) & "'," & datForm(Now) & ")")
      Set rs = DBCn.Execute("select * from laborxeingel where name = '" & SL.Item(ii).File.Name & "' and pfad = '" & umw(SL.Item(ii).File.Path) & "'")
      SL.Item(ii).DatID = rs!DatID
    End If
'   If f1.Name Like "1b*.ld*" Or f1.Name Like "Labor*.dat" Then
'   If f1.Name Like "Labor*.dat" Then
'    repl = Replace(f1.Name, ".", "")
'    Dim DatID&
'    DatID = rs!DatID
     
     Open SL.Item(ii).File.Path For Input As #1
     
  Dim obEOF%, obAngeh%, obSchluss%, ltxt$
'  Dim rLsAkt As New ADODB.Recordset, rLuAkt As New ADODB.Recordset
  obSchluss = 0
  obEOF = EOF(1)
  Do
   If obEOF Then Exit Do
   Line Input #1, txt
   frm.Zeilen = frm.Zeilen + 1
   frm.Zeilen = Format(frm.Bytes + Len(txt) + 2, "###,###,###,###,###,###,###,##0")
   If frm.Zeilen Mod 10 = 0 Then
    frm.Sekunden = Format((Now - t1) * 60 * 60 * 24, "###,###,###,###,###,###,##0")
    DoEvents
   End If
   obEOF = EOF(1)
   If Not Left(txt, 7) Like "#######" Then 'IsNumeric(Left(txt, 7)) Then
    ltxt = ltxt + " " + txt
    obAngeh = -1
   ElseIf obAngeh Then
    obAngeh = 0
   End If
  
   Do
    If obAngeh Then
     If obEOF Then
      obSchluss = True
     Else
      Exit Do
     End If
    End If
    If ZS Then ltxt = ZSU1(ltxt)
    If ltxt <> "" Then
'    Call dolies(ltxt, frm, obSchluss) ' hierin auch alleSpeichern
     
'     Do
      Dim Lenge%, Kennung%, Inhalt$
      
      absPos = absPos + 1
      Lenge = Left(ltxt, 3)
      Kennung = Mid(ltxt, 4, 4)
      Inhalt = Replace(umw(Mid(ltxt, 8)), "³", "ü") ' Zeichensatz( ist meist falsch, nur das ü stimmt in Karlsfeld nicht
      DebugZähler = DebugZähler + 1
        Select Case Kennung
         Case 8000
          Select Case Left(Inhalt, 4)
           Case 8220 'L-Datenpaket-Header (Turbomed)
            ReDim rLs(0)
            ReDim rLu(0)
            ReDim rLo(0)
            ReDim rLw(0)
            ReDim rLL(0)
            rLs(0).Satzart = Inhalt
            RefNr = 0
            SatzlängenArt = 0
            obSonderSatz = -1
           Case 8221 ' Schluss ist erst bei 9202
            SatzlängenArt = -1
            obSonderSatz = -1
           Case Else ' 8201, 8202,8203
'            If rLuAkt.State = 1 Then rLuAkt.Update
            ReDim Preserve rLu(UBound(rLu) + 1)
            RefNr = RefNr - 1 ' negative Pseudo-Refnummern
            rLu(UBound(rLu)).DatID = SL.Item(ii).DatID
            rLs(0).DatID = SL.Item(ii).DatID
            rLu(UBound(rLu)).SatzID = rLs(0).SatzID
            rLu(UBound(rLu)).Satzart = Inhalt
            rLu(UBound(rLu)).RefNr = RefNr
            obSonderSatz = 0
          End Select
         Case 8410
          ReDim Preserve rLw(UBound(rLw) + 1)
          rLw(UBound(rLw)).RefNr = rLu(UBound(rLu)).RefNr
          obB = 0
          rLw(UBound(rLw)).Abkü = Inhalt 'Replace(Inhalt, "€", " ")
         Case 8434
          ReDim Preserve rLo(UBound(rLo) + 1)
          rLo(UBound(rLo)).RefNr = rLu(UBound(rLu)).RefNr
          obB = -1
          rLo(UBound(rLo)).Verf = Inhalt
         Case 8100: If obSonderSatz Then If Not SatzlängenArt Then rLs(0).Satzlänge = Inhalt Else rLs(0).SatzlängeSchluss = Inhalt Else rLu(UBound(rLu)).Satzlänge = Inhalt
         Case 8310: rLu(UBound(rLu)).Auftragsnummer = Inhalt
         Case 8311: rLu(UBound(rLu)).Auftragsschlüssel = Inhalt
         Case 8301: rLu(UBound(rLu)).Eingang = BDTtoDate(Inhalt)
         Case 8302: rLu(UBound(rLu)).Berichtsdatum = Inhalt
         Case 8303: rLu(UBound(rLu)).Berichtsdatum = rLu(UBound(rLu)).Berichtsdatum & " " & Inhalt
         Case 8609: rLu(UBound(rLu)).Abrechnungstyp = Inhalt
         Case 8401: rLu(UBound(rLu)).BefArt = Inhalt ' Fertigstellungsgrad
         Case 8403: rLu(UBound(rLu)).GebüOrd = Inhalt
         Case 8405: rLu(UBound(rLu)).Patienteninformation = Inhalt
         Case 8407: rLu(UBound(rLu)).Geschlecht = Inhalt
         Case 3101:
          rLu(UBound(rLu)).Nachname = Inhalt
          If rLu(UBound(rLu)).Nachname Like "zzz*" Then rLu(UBound(rLu)).Nachname = Mid(rLu(UBound(rLu)).Nachname, 4)
         Case 3102: rLu(UBound(rLu)).Vorname = Inhalt
         Case 3103
          Dim glZ%, glZLocker%, glZoL%
'          Dim Bm$, bmLocker$, BmoL$
          Dim erwZ% 'Zahl der erwogenen Pat_ids
          Dim rLuGebDat As Date
          Dim merkPat_ID&
          erwZ = 0
          rLu(UBound(rLu)).GebDat = Inhalt
          If Left(rLu(UBound(rLu)).GebDat, 4) = "0000" Then rLu(UBound(rLu)).GebDat = "0101" + Right(rLu(UBound(rLu)).GebDat, 4)
          rLuGebDat = BDTtoDate(rLu(UBound(rLu)).GebDat) ' CDate(Format(rlu(ubound(rlu)).GebDat, "##/##/####"))
          If Not (IsNull(rLu(UBound(rLu)).Nachname) Or rLu(UBound(rLu)).Nachname = "") Or Not (IsNull(rLu(UBound(rLu)).Vorname) Or rLu(UBound(rLu)).Vorname = "") Then
           Dim Runde%
           Dim Nachname$
           For Runde = 1 To 2
'           Select Case runde
'            Case 1: rNam.Seek "=", rlu(ubound(rlu)).Nachname, rlu(ubound(rlu)).Vorname, rlu(ubound(rlu)).GebDat
'            Case 2: rNam.Seek "=", rlu(ubound(rlu)).Nachname, rlu(ubound(rlu)).Vorname
'           End Select
            If rs.State = 1 Then rs.Close
            Select Case Runde
             Case 1:
              rs.Open "select * from namen where nachname = '" & rLu(UBound(rLu)).Nachname & "' and vorname = '" & rLu(UBound(rLu)).Vorname & "' and gebdat = " & datForm(rLuGebDat), DBCn, adOpenStatic, adLockReadOnly
              If rs.BOF Then rs.Close: rs.Open "select * from namen where nachname = 'zzz" & rLu(UBound(rLu)).Nachname & "' and vorname = '" & rLu(UBound(rLu)).Vorname & "' and gebdat = " & datForm(rLuGebDat), DBCn, adOpenStatic, adLockReadOnly
             Case 2: rs.Open "select * from namen where nachname = '" & rLu(UBound(rLu)).Nachname & "' and vorname = '" & rLu(UBound(rLu)).Vorname & "'", DBCn, adOpenStatic, adLockReadOnly
              If rs.BOF Then rs.Close: rs.Open "select * from namen where nachname = 'zzz" & rLu(UBound(rLu)).Nachname & "' and vorname = '" & rLu(UBound(rLu)).Vorname & "'", DBCn, adOpenStatic, adLockReadOnly
            End Select
            If Not rs.BOF Then
             Do
              If Runde = 1 Or rs!GebDat <> rLuGebDat Then  ' damit in der zweiten Runde nicht dieselben nochmal kommen / Null dürfte nicht vorkommen
               erwZ = erwZ + 1
               Select Case Runde
                Case 1
                 rLu(UBound(rLu)).Pat_idErwVNG = rLu(UBound(rLu)).Pat_idErwVNG & IIf(IsNull(rLu(UBound(rLu)).Pat_idErwVNG) Or rLu(UBound(rLu)).Pat_idErwVNG = "", "", "/") & CStr(rs!Pat_ID)
                Case 2
                 rLu(UBound(rLu)).Pat_idErwVN = rLu(UBound(rLu)).Pat_idErwVN & IIf(IsNull(rLu(UBound(rLu)).Pat_idErwVN) Or rLu(UBound(rLu)).Pat_idErwVN = "", "", "/") & CStr(rs!Pat_ID)
               End Select
               merkPat_ID = rs!Pat_ID
              End If
              rs.Move 1
              If rs.EOF Then Exit Do
              Nachname = rs!Nachname
              If Nachname Like "zzz*" Then Nachname = Mid(Nachname, 4)
              If Nachname <> rLu(UBound(rLu)).Nachname Or rs!Vorname <> rLu(UBound(rLu)).Vorname Or (Runde = 1 And rs!GebDat <> rLuGebDat) Then Exit Do
             Loop
            End If ' not nomatch
           Next Runde
          End If ' isnull...
          If erwZ = 1 Then
           rLu(UBound(rLu)).Pat_ID = merkPat_ID
           rLu(UBound(rLu)).Pat_idUrsp = "E"
          Else
           glZ = 0 ' Zahl der Patienten mit passendem Geburtstag, zu denen ein zeitlich passendes Labor in den anderen Labortabellen gefunden wurde
           glZLocker = 0 ' Zahl der Patienten mit passendem Geburtstag, die zum Zeitpunkt der Laboruntersuchung als schon bei mir behandelt gekennzeichnet waren
           glZoL = 0 ' Zahl der Patienten mit passendem Geburtstag
'           Bm = ""
'           bmLocker = ""
'           BmoL = ""
           If rs.State = 1 Then rs.Close
           rs.Open "select * from anamnesebogen where gebdat = " & datForm(rLuGebDat)
           If Not rs.BOF Then
            Do
             glZoL = glZoL + 1
             erwZ = erwZ + 1
             rLu(UBound(rLu)).Pat_idErwG = rLu(UBound(rLu)).Pat_idErwG & IIf(IsNull(rLu(UBound(rLu)).Pat_idErwG) Or rLu(UBound(rLu)).Pat_idErwG = "", "", "/") & CStr(rs!Pat_ID)
'             BmoL = rAnam.Bookmark
             If rs!Vorgestellt <= rLu(UBound(rLu)).Eingang Then
'              bmLocker = rAnam.Bookmark
              glZLocker = glZLocker + 1
              If glZLocker = 1 Then
                rLu(UBound(rLu)).Pat_ID = rs!Pat_ID
                rLu(UBound(rLu)).Pat_idUrsp = "E"
              End If
              rLu(UBound(rLu)).Pat_idErwGB = rLu(UBound(rLu)).Pat_idErwGB & IIf(IsNull(rLu(UBound(rLu)).Pat_idErwGB) Or rLu(UBound(rLu)).Pat_idErwGB = "", "", "/") & rs!Pat_ID
              If rs1.State = 1 Then rs1.Close
              rs1.Open "select * from laborneu where pat_id = " & rs!Pat_ID & " and zeitpunkt >= " & datForm(rLu(UBound(rLu)).Eingang - 5) & " and zeitpunkt <= " & datForm(rLu(UBound(rLu)).Eingang + 15), DBCn, adOpenStatic, adLockReadOnly
              If Not rs1.BOF Then
 '              Bm = rAnam.Bookmark
               glZ = glZ + 1
               rLu(UBound(rLu)).Pat_idErwGL = rLu(UBound(rLu)).Pat_idErwGL & IIf(IsNull(rLu(UBound(rLu)).Pat_idErwGL) Or rLu(UBound(rLu)).Pat_idErwGL = "", "", "/") & rs!Pat_ID
              End If
             End If
             rs.Move 1
             If rs.EOF Then Exit Do
            Loop
           End If
          End If ' erwZ = 1 nach Suche in Namen
         Case 3104: rLu(UBound(rLu)).Titel = Inhalt
         Case 8411: rLw(UBound(rLw)).Langname = Inhalt
         Case 5001
          ReDim Preserve rLL(UBound(rLL) + 1)
          rLL(UBound(rLL)).RefNr = rLu(UBound(rLu)).RefNr
          If obB Then rLL(UBound(rLL)).Verf = rLo(UBound(rLo)).Verf Else rLL(UBound(rLL)).Abkü = rLw(UBound(rLw)).Abkü
          rLL(UBound(rLL)).EBM = Inhalt
         Case 5005
          rLL(UBound(rLL)).Anzahl = Inhalt
         Case 8406
        ' zwingende Abfolge mit 5001 wurde 2/05 geprüft
          rLL(UBound(rLL)).goä = Inhalt
         Case 8428: rLo(UBound(rLo)).KuQu = Inhalt
         Case 8430: If obB Then rLo(UBound(rLo)).Quelle = Inhalt Else rLw(UBound(rLw)).Quelle = Inhalt
         Case 8431: If obB Then rLo(UBound(rLo)).Quelle = Inhalt Else rLw(UBound(rLw)).Quelle = Inhalt
         Case 8432: If obB Then rLo(UBound(rLo)).AbnDat = BDTtoDate(Inhalt) Else rLw(UBound(rLw)).AbnDat = BDTtoDate(Inhalt)
         Case 8418: rLw(UBound(rLw)).Teststatus = Inhalt
         Case 8420: rLw(UBound(rLw)).Wert = Inhalt
         Case 8421: rLw(UBound(rLw)).Einheit = Inhalt
         Case 8422: rLw(UBound(rLw)).Grenzwerti = Inhalt
         Case 8480
           If obB Then
            rLo(UBound(rLo)).Kommentar = IIf(IsNull(rLo(UBound(rLo)).Kommentar), "", rLo(UBound(rLo)).Kommentar + Chr(13) + Chr(10)) + Inhalt
            If keimz Then
             rLo(UBound(rLo)).Keimzahl = Inhalt
             keimz = 0
            End If
            If InStr(Inhalt, "Keimzahl") > 0 Then keimz = -1
           Else
            rLw(UBound(rLw)).Kommentar = nz(rLw(UBound(rLw)).Kommentar,"") + Inhalt
           End If
         Case 8490:
'          rLU!auftrhinw = inhalt ' weiß noch nicht, auf welcher Ebene relevant
           rLw(UBound(rLw)).AuftrHinw = Inhalt
         Case 8470
          If obB Then
           rLo(UBound(rLo)).Erklärung = IIf(IsNull(rLo(UBound(rLo)).Erklärung), "", rLo(UBound(rLo)).Erklärung + Chr(13) + Chr(10)) + Inhalt
          Else
           rLw(UBound(rLw)).Erklärung = nz(rLw(UBound(rLw)).Erklärung,"") + Inhalt
          End If
         Case 8460: rLw(UBound(rLw)).Normbereich = Inhalt
         Case 8461: rLw(UBound(rLw)).NormU = Inhalt
         Case 8462: rLw(UBound(rLw)).NormO = Inhalt
         Case 9212: rLs(0).VersionSatzb = Inhalt
         Case 201: rLs(0).Arztnr = Inhalt
         Case 203: rLs(0).Arztname = Inhalt
         Case 205: rLs(0).StraßePraxis = Inhalt
         Case 215: rLs(0).PLZPraxis = Inhalt
         Case 216: rLs(0).OrtPraxis = Inhalt
         Case 8300:
          Dim Spli$()
          Spli = Split(Inhalt, ";")
          If UBound(Spli) > 1 Then
           rLs(0).Labor = Spli(0)
           rLs(0).StraßeLabor = Spli(1)
           rLs(0).PLZLabor = Left(Spli(2), 5)
           rLs(0).OrtLabor = Mid(Spli(2), 6)
          End If
         Case 8320: rLs(0).Labor = Inhalt
         Case 8321: rLs(0).StraßeLabor = Inhalt
         Case 8322: rLs(0).PLZLabor = Inhalt
         Case 8323: rLs(0).OrtLabor = Inhalt
         Case 101: rLs(0).KBVPrüfnr = Inhalt
         Case 9106: rLs(0).Zeichensatz = Inhalt
         Case 8312: rLs(0).Kundenarztnr = Inhalt
         Case 9103: rLs(0).Erstellungsdatum = Inhalt
         Case 9202: rLs(0).Gesamtlänge = Inhalt
          If Not debugohneSpeichern Then
           Call laborxsaetzeSpeichern
          End If
          If rs.State = 1 Then rs.Close
          rs.Open "select max(satzid) as msatzid from laborxsaetze;"
          For i = 1 To UBound(rLu)
           rLu(i).SatzID = nz(rs!msatzid,0)
           If Not debugohneSpeichern Then
            Call laborxusSpeichern(i)
           End If
           If rs1.State = 1 Then rs1.Close
           rs1.Open "select max(refnr) as mrefnr from laborxus;", DBCn, adOpenStatic, adLockReadOnly
           Dim rs1refnr&
           If IsNull(rs1!mrefnr) Then rs1refnr = 0 Else rs1refnr = rs1!mrefnr
           For j = 1 To UBound(rLo)
            If rLo(j).RefNr = rLu(i).RefNr Then
             rLo(j).RefNr = rs1refnr
            End If
           Next j
           For j = 1 To UBound(rLw)
            If rLw(j).RefNr = rLu(i).RefNr Then
             rLw(j).RefNr = rs1refnr
            End If
           Next j
           For j = 1 To UBound(rLL)
            If rLL(j).RefNr = rLu(i).RefNr Then
             rLL(j).RefNr = rs1refnr
            End If
           Next j
          Next i
           If Not debugohneSpeichern Then
            Call laborxbaktSpeichern
            Call laborxwertSpeichern
            Call laborxleistSpeichern
           End If
         Case Else
          frm.Ausgabe = "Unbekannte Kennung: " & Kennung & ", Inhalt: " & Inhalt & vbCrLf & frm.Ausgabe
        End Select
'      Loop
     
     If obSchluss Or BrichAb Then Exit Do
    End If
    ltxt = txt
    If Not obEOF Then Exit Do
    obSchluss = True
   Loop
   If BrichAb Then Exit Do
  Loop
  Close #1
           
'    End If 'Not (oballe Or obDateiNeu) Then ' oballe or not FSO.fileexists(neupfad)
    frm.Bytes = frm.Bytes + 1
    If frm.Bytes Mod 100 = 0 Then
     frm.Prozent = Format(frm.Bytes / frm.GesBytes * 100, "0.00")
     frm.Sekunden = Format((Now - t1) * 60 * 60 * 24, "###,###,###,###,###,###,##0")
     frm.Beginn = Format(t1, "hh:mm:ss")
     frm.GesDauer = Format((t1a - t1 + frm.GesBytes / frm.Bytes * (Now - t1a)), "hh:mm:ss")
     frm.Ende = Format(t1 + (t1a - t1 + frm.GesBytes / frm.Bytes * (Now - t1a)), "hh:mm:ss")
   
    End If
    DoEvents
    If BrichAb Then
     frm.Ausgabe = "Einlesen der Labor-DFÜ-Dateien abgebrochen " + vbCrLf + frm.Ausgabe
     Exit Function
    End If
  Next
  frm.Ausgabe = "Fertig mit Einlesen der Labor-DFÜ-Dateien" + vbCrLf + frm.Ausgabe
  frm.ZeilenBez = "Zeilen:"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborDirektImport/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LaborDirektImport
Function LaborErgPatId(frm As Lese, Optional mitLöschen%, Optional nurRef&)
' Pat_IDs ergänzen
 Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
 Const debugbit% = 0
 Dim altRefnr&, i&, sqlct$
 Dim j&
 Dim sPI As SortierPat_ID
 Dim SL As New SortierListe
' Dim rLX As DAO.Recordset, rLN As DAO.Recordset, rs As DAO.Recordset
 Dim rLX As New ADODB.Recordset, rLN As New ADODB.Recordset
 Dim rLNZeitpunkt#, ZdüP%, LWerte$, GesAbk$, raf&, laborneuAfn&, sql$, sql1$
 Dim obEintragen% ' Datensatz in rLX eintragen und mit SL neu anfangen
 Dim aktDS&, VergleichTot%
 
' nurRef = 3799
 On Error GoTo fehler
 t1a = Now
 frm.BytesBez = "Zeilen"
 frm.ZeilenBez = ""
 
 GesAbk = "("
 Dim altAusgabe$
 altAusgabe = frm.Ausgabe
 frm.Ausgabe = "Vergleiche die Laborübertragungsdateien mit dem BDT-Export ..." + vbCrLf + frm.Ausgabe
 frm.BytesBez = "Sätze:"
 frm.Bytes = 0
' If DBCn Is Nothing Or DBCn = "" Then Call ConstrFestleg(0, frm.dlg)
  '"ZdüP, LWerte "
 ' mitLöschen = -1
  If mitLöschen Then
   If Not debugbit Then
    Call löschBezügeausLaborux(nurRef)
   End If
  End If
  
  sql = "SELECT distinct Auftragsnummer, Eingang, LaborXUS.Pat_id, Nachname, Vorname, GebDat, Geschlecht, GebüOrd, " & _
        "Erstellungsdatum, Wert, Einheit, LaborXWert.Abkü, LaborXUS.RefNr, Pat_idUrsp, Pat_idLaborNeu " & _
        "FROM (LaborXSaetze INNER JOIN (LaborXEingel INNER JOIN LaborXUS ON LaborXEingel.DatID = LaborXUS.DatID) " & _
        "ON LaborXSaetze.SatzID = LaborXUS.SatzID) INNER JOIN LaborXWert ON LaborXUS.RefNr = LaborXWert.RefNr " & _
        "WHERE not isnull(LaborXWert.Wert) and laborxwert.wert <> """" and isnull(verglichen) "
' die leeren werden oft gar nicht in der BDT-Datei mitgeliefert, was zu hohen Raten unangemessener Aberkennungen führt
'  If Not IsNull(nurRef) Then
   If nurRef > 0 Then
    sql = sql + " AND laborxus.Refnr in (" & nurRef & ") "
   End If
'  End If
'  sql = sql + " ORDER BY eingang, laborxus.refnr"
   sql = sql + " ORDER BY laborxus.refnr"
   
  sqlct = "select count(*) as ct from (" & sql & ") as Zl"
  Set rs = DBCn.Execute(LCase(sqlct))
  frm.GesBytes = rs!ct
   
   If mitLöschen Then
'   If rs!ct > 30000 Then
    If Not debugbit Then
     Call LöschRefNr(nurRef)
    End If
'   End If
  End If

  If debugbit Then Open "u:\anamnese\debug.txt" For Output As #300
' Set rLX = dbcn.Execute(LCase(sql))
 obEintragen = -1 ' SL neu befüllen
 rLX.Open LCase(sql), DBCn, adOpenDynamic, adLockOptimistic
 Do While Not rLX.EOF
  aktDS = aktDS + 1
'  SysCmd 4, "Vergleiche Labore: " + CStr(aktDS) + "/" + CStr(DSZahl) ' rs!ct von oben
' Falls in den entsprechenden Feldern von rLX schon was steht => löschen
'  rLN.Open "laborneu", dbcn
  ZdüP = ZdüP + 1 ' Zahl der überprüften Parameter
  LWerte = LWerte + rLX!Abkü + ":" + rLX!Wert + ", "
  If Not VergleichTot Then ' solange der Vergleich noch tot ist ...
'   Print #300, vbCrLf & "Refnr: " & rLX!RefNr & " (Pat_idLaborneu: " & rLX!Pat_idLaborNeu & ": LWerte: " & LWerte
   If rLN.State = 1 Then rLN.Close
   rLN.Open "select distinct pat_id, abkü, wert, zeitpunkt from laborneu where abkü = '" & rLX!Abkü & "' and wert = '" & rLX!Wert & "'", DBCn, adOpenDynamic, adLockOptimistic
   Dim debugzl&
   If debugbit Then
    debugzl = 0
    Do
     If rLN.BOF Or rLN.EOF Then Exit Do
      debugzl = debugzl + 1
      Print #300, "  " & debugzl & ": rLN!Pat_id: " & rLN!Pat_ID, rLN!Abkü, rLN!Wert, rLN!Zeitpunkt
     rLN.Move 1
    Loop
    If Not rLN.BOF Then
     rLN.MoveFirst
    End If
   End If
   If Not rLN.BOF Then
    GesAbk = GesAbk & "'" & rLN!Abkü & "',"
    If obEintragen Then
     Do
      If Abs(rLN!Zeitpunkt - rLX!Eingang) < 32 Then ' 1 Monat Grenze
' SL = Liste der Patienten, die für die aktuelle Laborzusammenstellung in Frage kommen
       Set sPI = New SortierPat_ID
       sPI.Pat_ID = rLN!Pat_ID
       sPI.Zeitpunkt = rLN!Zeitpunkt
       Call SL.Add(sPI)
      End If
      rLN.MoveNext
      If rLN.EOF Then Exit Do
'      If rLN!Abkü <> rLX!Abkü Or rLN!Wert <> rLX!Wert Then Exit Do
     Loop
    Else ' obeintragen
' diejenigen der bisher in Frage kommenden Patienten, die auch für den aktuellen Laborwert noch in Frage kommen, erneut kennzeichnen
'     If rLN!Abkü = "CRP" Then Stop ' And rLN!Pat_id = 412
     Do
'      For Each sPI In SL
      For i = 1 To SL.Count
       Set sPI = SL.Item(i)
       If sPI.Pat_ID = rLN!Pat_ID And sPI.Zeitpunkt = rLN!Zeitpunkt Then
        sPI.Knz = -1
        Exit For
       End If
      Next i
      rLN.MoveNext
      If rLN.EOF Then Exit Do
      If UCase(rLN!Abkü) <> UCase(rLX!Abkü) Or rLN!Wert <> rLX!Wert Then
       Exit Do
      End If
     Loop
' die restlichen löschen und die Kennzeichen zurücksetzen
     For i = SL.Count To 1 Step -1
      Set sPI = SL.Item(i)
      If sPI.Knz = 0 Then
       Call SL.Remove(i)
      Else
       sPI.Knz = 0
      End If
     Next i
     If SL.Count = 0 Then VergleichTot = -1
     If debugbit Then
      Print #300, " SL.Count: " & SL.Count
      For i = 1 To SL.Count
       Print #300, "   Pat_id: " & SL.Item(i).Pat_ID
      Next i
     End If
    End If ' obeintragen
   End If ' Not rLN.NoMatch Then
  End If ' VergleichTot
  altRefnr = rLX!RefNr
  rLX.Move 1
'  If rLX.EOF Then obEintragen = -1 Else If rLX!Auftragsnummer <> altAuftragsNummer Then obEintragen = -1 Else obEintragen = 0
  If rLX.EOF Then obEintragen = -1 Else If rLX!RefNr <> altRefnr Then obEintragen = -1 Else obEintragen = 0
  If obEintragen Then
' dann fängt später der nächste Auftrag an, der aktuelle wird eingetragen
    Dim nPat_idLaborNeu$, nPat_id&, nPat_idUrsp$, SLPat_id&
    VergleichTot = 0
    rLX.Move -1
    
    SLPat_id = -1
    If SL.Count = 0 Then
    ElseIf SL.Count = 1 Then
     SLPat_id = SL.Item(1).Pat_ID
     rLNZeitpunkt = SL.Item(1).Zeitpunkt
     nPat_idLaborNeu = SLPat_id
    ElseIf rLX!Pat_ID <> 0 And Not IsNull(rLX!Pat_ID) Then ' wenn die Laborwertekombination für mehrere Patienten stimmen würde, aber der richtige schon aus Namen/Geburtsdatum hervorgeht => diesen davon nehmen
     For i = 1 To SL.Count
      If SL.Item(i).Pat_ID = rLX!Pat_ID Then
       SLPat_id = SL.Item(i).Pat_ID
       rLNZeitpunkt = SL.Item(i).Zeitpunkt
       nPat_idLaborNeu = SLPat_id
       Exit For
      End If
     Next
    Else ' wenn also unbekannt ist, welcher Patient der richtige ist
     nPat_idLaborNeu = SL.Item(1).Pat_ID
     For i = 2 To SL.Count
      nPat_idLaborNeu = nPat_idLaborNeu & "/" & SL.Item(i).Pat_ID
     Next
    End If
    
    If SLPat_id <> -1 Then ' SL.Count = 1
     nPat_id = nz(rLX!Pat_ID,0)
     nPat_idUrsp = nz(rLX!Pat_idUrsp,"")
'    If SL.Count = 1 Then
     If nPat_id = 0 Then
      nPat_id = SLPat_id
      nPat_idUrsp = "L"
     ElseIf nPat_id = SLPat_id Then
      nPat_idUrsp = "B" ' "E"rsteinlesung, "B"eide, "L"abordirektdateien
     Else
      nPat_idUrsp = "W" ' "W"idersprüchliche Patientenangaben!
     End If
     
     GesAbk = Left(GesAbk, Len(GesAbk) - 1) + ")"
     If Not debugbit Then
      Call DBCn.Execute("update laborneu set refnr = " & rLX!RefNr & " where pat_id = " & SLPat_id & " and zeitpunkt = " & datForm(rLNZeitpunkt) & " and abkü in " & GesAbk, laborneuAfn)
      If laborneuAfn = 0 Then Err.Raise 9999, , "LaborneuAFN = 0"
     End If
'    End If ' SL.count = 1
    
     If Not debugbit Then
      Call DBCn.Execute("update laborxus set pat_idUrsp= '" & nPat_idUrsp & "',Pat_id = " & nPat_id & ",ZeitpunktLaborneu = " & datForm(rLNZeitpunkt) & ",LWerte = '" & LWerte & "',ZdüP = " & ZdüP & ", Pat_idLaborNeu = '" & nPat_idLaborNeu & "', AfN = " & laborneuAfn & " where refnr = " & rLX!RefNr, raf)
     End If
    End If ' SL.count = 1
    If SLPat_id <> -1 Or (Now - BDTtoDate(rLX!Erstellungsdatum)) > 7 Then
     If Not debugbit Then
      Call DBCn.Execute("update laborxus set verglichen = " & datForm(Now) & ",LWerte = '" & LWerte & "',ZdüP = " & ZdüP & ", ZdiP = " & SL.Count & " where refnr = " & rLX!RefNr, raf)
     End If
    End If
    rLX.Move 1 ' nach vorher -1
'   End If '    If SL.Count <= 1 Or IsNull(rLX!Pat_ID) Or rLX!Pat_ID = 0 Or Not obEnthalten Then
     For i = SL.Count To 1 Step -1
      Call SL.Remove(i)
     Next i
     ZdüP = 0
     GesAbk = "("
     LWerte = ""
  End If ' obeintragen
  frm.Bytes = frm.Bytes + 1
  If frm.Bytes Mod 100 = 0 Then
   If frm.GesBytes <> 0 Then frm.Prozent = Format(frm.Bytes / frm.GesBytes * 100, "0.00")
   frm.Sekunden = Format((Now - t1) * 60 * 60 * 24, "###,###,###,###,###,###,##0")
   frm.Beginn = Format(t1, "hh:mm:ss")
   DoEvents
  End If
  
  If rLX.EOF Then Exit Do
  If BrichAb Then
   frm.Ausgabe = "LaborErgPatID (Vergleich der Labor-DFÜ-Datein mit dem BDT-Export) bei RefNr " & rLX!RefNr & " abgebrochen" & vbCrLf & altAusgabe
   Exit Function
  End If
 Loop ' while not rlx.eof
 frm.Ausgabe = "Fertig mit LaborErgPatId (Vergleich der Labor-DFÜ-Datein mit dem BDT-Export)" & vbCrLf & altAusgabe
 If debugbit Then
  Close #300
 End If
 frm.BytesBez = "Bytes:"
 frm.ZeilenBez = "Zeilen:"
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LaborErgPatID/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LaborErgPatID

Function LöschRefNr(Optional RefNr&)
 Dim raf&
 If RefNr = 0 Then
  Set rs = DBCn.Execute("update laborneu set refnr = null where not isnull(refnr)", raf)
 Else
  Set rs = DBCn.Execute("update laborneu set refnr = null where refnr = " & RefNr, raf)
 End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LöschRefNr/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LöschRefNr
#If False Then
Function LöschRefNr()
  On Error GoTo fehler
  Call dtbInit
#If True Then
 Dtb.Execute ("update laborneu set refnr = null")
#Else
  Dim rs As DAO.Recordset
  Set rs = TabÖff("LaborNeu")
  Do While Not rs.EOF
   If Not IsNull(rs!RefNr) Then
    rs.Edit
    rs!RefNr = Null
    rs.Update
   End If
   rs.MoveNext
  Loop
#End If
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in LöschRefNr/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' LöschRefNr
#End If
Function testLöschAb(ab&)
 Call ConstrFestleg(0, Dialog)
' dbcn.Open ConStr
 Call DBCn.Execute("delete from laborxleist where refnr in (select refnr from laborxus where datid>=" & ab & ");")
 Call DBCn.Execute("delete from laborxbakt where refnr in (select refnr from laborxus where datid>=" & ab & ");")
 Call DBCn.Execute("set foreign_key_checks=0")
 Call DBCn.Execute("delete from laborxwert where refnr in (select refnr from laborxus where datid>=" & ab & ");")
 Call DBCn.Execute("set foreign_key_checks=1")
 Call DBCn.Execute("delete from laborxus where datid >= " & ab & ";")
 Call DBCn.Execute("delete from laborxeingel where datid >= " & ab & ";")
End Function ' testLöschab
Function EmailsImport()
 Dim con As New ADODB.Connection  ' Connection
 Dim rNa As New ADODB.Recordset
 Dim rEx As New ADODB.Recordset
 Dim rx As New ADOX.Catalog
 Const EmDatei$ = "p:\Patientenübergreifendes\Emails.xls" ' Excel-Datei mit Suche aus Turbomed "*@*"
 On Error GoTo fehler
 If DBCn Is Nothing Or DBCn = "" Then Call ConstrFestleg(0, Dialog)
 con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=""Excel 8.0;HDR=No;IMEX=2"";Data Source=" & EmDatei & ";" ' TABLE=Adressen$"
 Dim Runde%, i%, pFeld$, eFeld$, obAnfang%, pNr%, pRoh$, Email$, ka%, ke%, raf&
 rx.ActiveConnection = con
 rEx.Open "[" & rx.Tables(rx.Tables.Count - 1).Name & "]", con
 Do While Not rEx.EOF
'  Debug.Print runde
  If obAnfang Then
    pRoh = rEx.Fields(pFeld)
    ka = InStr(pRoh, "(")
    ke = InStr(pRoh, ")")
    If ka > 0 And ke > 0 Then
     pNr = Val(Mid(pRoh, ka + 1, ke - 1))
     Call DBCn.Execute("update namen set email = '" & rEx.Fields(eFeld) & "' where pat_id = " & pNr, raf)
     If raf <> 1 Then Err.Raise 999, , "Fehler beim Einfügen der Email-Adresse '" & rEx.Fields(eFeld) & "' zum Patienten Nr. " & pNr
    End If
  ElseIf Not IsNull(rEx.Fields(0)) And Not IsNull(rEx.Fields(1)) Then
   For i = 0 To rEx.Fields.Count - 1
    If rEx.Fields(i) = "Patient" Then
      pFeld = rEx.Fields(i).Name
    ElseIf rEx.Fields(i) = "Email" Then
      eFeld = rEx.Fields(i).Name
    End If
   Next i
   obAnfang = True
  End If
  Runde = Runde + 1
  rEx.MoveNext
 Loop
 On Error Resume Next
 rEx.Close
 Exit Function
fehler:
 Dim AnwPfad$
#If VBA6 Then
 AnwPfad = CurrentDb.Name
#Else
 AnwPfad = App.Path
#End If
Select Case MsgBox("FNr: " + CStr(Err.Number) + vbCrLf + "LastDLLError: " + CStr(Err.LastDllError) + vbCrLf + "Source: " + IIf(IsNull(Err.Source), "", CStr(Err.Source)) + vbCrLf + "Description: " + Err.Description, vbAbortRetryIgnore, "Aufgefangener Fehler in EmailsImport/" + AnwPfad)
 Case vbAbort: Call MsgBox("Höre auf"): End
 Case vbRetry: Call MsgBox("Versuche nochmal"): Resume
 Case vbIgnore: Call MsgBox("Setze fort"): Resume Next
End Select
End Function ' Emails-Import

